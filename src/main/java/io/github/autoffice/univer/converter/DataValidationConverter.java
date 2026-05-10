package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.UUID;

/**
 * 数据验证转换器。
 * Data validation converter between POI XSSF validation rules and Univer plugin resources.
 * <p>
 * 读取时把每个 sheet 的 {@link DataValidation} 转成 {@code SHEET_DATA_VALIDATION_PLUGIN}
 * 的 per-sheet 资源；写入时把该资源还原为 POI validation 规则。
 */
public final class DataValidationConverter {

    /** Univer 数据验证插件名 / Univer data validation plugin resource name. */
    public static final String PLUGIN_NAME = "SHEET_DATA_VALIDATION_PLUGIN";

    private DataValidationConverter() {
    }

    /**
     * 读取 sheet 中全部数据验证规则，转为按 ruleId 分组的资源 JSON。
     * Read all data validation rules from a sheet into a per-rule JSON object.
     */
    public static ObjectNode readSheetDataValidations(XSSFSheet sheet, ObjectMapper mapper) {
        ObjectNode out = mapper.createObjectNode();
        if (sheet == null) {
            return out;
        }
        for (DataValidation validation : sheet.getDataValidations()) {
            ObjectNode rule = toRuleJson(validation, mapper);
            if (rule != null) {
                out.set(rule.path("uid").asText(), rule);
            }
        }
        return out;
    }

    /**
     * 把 Univer 数据验证资源写回 sheet。
     * Write Univer data validation resource back onto the POI sheet.
     */
    public static void writeSheetDataValidations(XSSFSheet sheet, JsonNode rulesNode) throws io.github.autoffice.univer.UniverXlsxUnsupportedFeatureException {
        writeSheetDataValidations(sheet, rulesNode, null);
    }

    /**
     * 把 Univer 数据验证资源写回 sheet（带选项）。
     * Write Univer data validation resource back onto the POI sheet with options.
     */
    public static void writeSheetDataValidations(XSSFSheet sheet, JsonNode rulesNode, io.github.autoffice.univer.UniverXlsxOptions opts) throws io.github.autoffice.univer.UniverXlsxUnsupportedFeatureException {
        if (sheet == null || rulesNode == null || !rulesNode.isObject() || rulesNode.size() == 0) {
            return;
        }
        DataValidationHelper helper = new XSSFDataValidationHelper(sheet);
        for (Map.Entry<String, JsonNode> entry : iterable(rulesNode.fields())) {
            JsonNode ruleNode = entry.getValue();
            if (ruleNode == null || !ruleNode.isObject()) {
                continue;
            }
            applyRule(sheet, helper, ruleNode, opts);
        }
    }

    private static ObjectNode toRuleJson(DataValidation validation, ObjectMapper mapper) {
        if (validation == null) {
            return null;
        }
        DataValidationConstraint constraint = validation.getValidationConstraint();
        if (constraint == null) {
            return null;
        }
        ObjectNode out = mapper.createObjectNode();
        out.put("uid", UUID.randomUUID().toString());
        out.set("ranges", rangesToJson(validation.getRegions(), mapper));
        out.put("type", mapType(constraint.getValidationType()));

        String operator = mapOperator(constraint.getOperator());
        if (operator != null) {
            out.put("operator", operator);
        }
        String formula1 = extractFormula1(constraint);
        String formula2 = normalizeFormulaForRead(constraint, constraint.getFormula2());
        if (formula1 != null) {
            out.put("formula1", formula1);
        }
        if (formula2 != null) {
            out.put("formula2", formula2);
        }
        out.put("allowBlank", validation.getEmptyCellAllowed());
        out.put("showDropDown", !validation.getSuppressDropDownArrow());
        out.put("showErrorMessage", validation.getShowErrorBox());
        out.put("showInputMessage", validation.getShowPromptBox());
        if (validation.getErrorBoxText() != null) {
            out.put("error", validation.getErrorBoxText());
        }
        if (validation.getErrorBoxTitle() != null) {
            out.put("errorTitle", validation.getErrorBoxTitle());
        }
        if (validation.getPromptBoxText() != null) {
            out.put("prompt", validation.getPromptBoxText());
        }
        if (validation.getPromptBoxTitle() != null) {
            out.put("promptTitle", validation.getPromptBoxTitle());
        }
        out.put("errorStyle", mapErrorStyle(validation.getErrorStyle()));
        return out;
    }

    private static void applyRule(XSSFSheet sheet, DataValidationHelper helper, JsonNode ruleNode, io.github.autoffice.univer.UniverXlsxOptions opts) throws io.github.autoffice.univer.UniverXlsxUnsupportedFeatureException {
        CellRangeAddressList ranges = jsonToRanges(ruleNode.path("ranges"));
        if (ranges.countRanges() == 0) {
            return;
        }
        String type = ruleNode.path("type").asText(null);
        String operator = ruleNode.path("operator").asText(null);
        String formula1 = ruleNode.path("formula1").asText(null);
        String formula2 = ruleNode.path("formula2").asText(null);
        DataValidationConstraint constraint = createConstraint(helper, type, operator, formula1, formula2, ruleNode);
        if (constraint == null) {
            boolean strict = opts != null && opts.isStrictMode();
            if (strict) {
                throw new io.github.autoffice.univer.UniverXlsxUnsupportedFeatureException(
                    "Unsupported or unknown data validation type: " + type);
            }
            System.err.println("WARNING: Skipping data validation rule with unsupported type: " + type);
            return;
        }
        DataValidation validation = helper.createValidation(constraint, ranges);
        validation.setEmptyCellAllowed(ruleNode.path("allowBlank").asBoolean(true));
        boolean showDropDown = ruleNode.path("showDropDown").asBoolean(false);
        validation.setSuppressDropDownArrow(!showDropDown);
        validation.setShowErrorBox(ruleNode.path("showErrorMessage").asBoolean(false));
        validation.setShowPromptBox(ruleNode.path("showInputMessage").asBoolean(false));
        String error = ruleNode.path("error").asText(null);
        String errorTitle = ruleNode.path("errorTitle").asText(null);
        String prompt = ruleNode.path("prompt").asText(null);
        String promptTitle = ruleNode.path("promptTitle").asText(null);
        if (error != null || errorTitle != null) {
            validation.createErrorBox(errorTitle == null ? "" : errorTitle, error == null ? "" : error);
        }
        if (prompt != null || promptTitle != null) {
            validation.createPromptBox(promptTitle == null ? "" : promptTitle, prompt == null ? "" : prompt);
        }
        validation.setErrorStyle(mapErrorStyleFromJson(ruleNode.path("errorStyle")));
        sheet.addValidationData(validation);
    }

    private static DataValidationConstraint createConstraint(DataValidationHelper helper,
                                                             String type,
                                                             String operator,
                                                             String formula1,
                                                             String formula2,
                                                             JsonNode ruleNode) {
        if (type == null) {
            return null;
        }
        switch (type) {
            case "whole":
                return createNumericConstraint(helper, true, operator, formula1, formula2);
            case "decimal":
                return createNumericConstraint(helper, false, operator, formula1, formula2);
            case "date":
                return createDateConstraint(helper, operator, formula1, formula2, ruleNode);
            case "textLength":
                return createTextLengthConstraint(helper, operator, formula1, formula2);
            case "list":
            case "listMultiple":
                return createListConstraint(helper, formula1, formula2, ruleNode.path("showDropDown").asBoolean(false));
            case "custom":
                return helper.createCustomConstraint(normalizeFormulaForWrite(formula1));
            case "any":
            case "none":
                return helper.createCustomConstraint("TRUE");
            default:
                return null;
        }
    }

    private static DataValidationConstraint createNumericConstraint(DataValidationHelper helper,
                                                                    boolean whole,
                                                                    String operator,
                                                                    String formula1,
                                                                    String formula2) {
        int op = mapOperatorToPoi(operator);
        String f1 = normalizeFormulaForWrite(formula1);
        String f2 = normalizeFormulaForWrite(formula2);
        if (whole) {
            if (op == DataValidationConstraint.OperatorType.IGNORED) {
                return helper.createIntegerConstraint(DataValidationConstraint.OperatorType.IGNORED, f1, f2);
            }
            return helper.createIntegerConstraint(op, f1, f2);
        }
        if (op == DataValidationConstraint.OperatorType.IGNORED) {
            return helper.createDecimalConstraint(DataValidationConstraint.OperatorType.IGNORED, f1, f2);
        }
        return helper.createDecimalConstraint(op, f1, f2);
    }

    private static DataValidationConstraint createTextLengthConstraint(DataValidationHelper helper,
                                                                       String operator,
                                                                       String formula1,
                                                                       String formula2) {
        int op = mapOperatorToPoi(operator);
        return helper.createTextLengthConstraint(op, normalizeFormulaForWrite(formula1), normalizeFormulaForWrite(formula2));
    }

    private static DataValidationConstraint createDateConstraint(DataValidationHelper helper,
                                                                  String operator,
                                                                  String formula1,
                                                                  String formula2,
                                                                  JsonNode ruleNode) {
        int op = mapOperatorToPoi(operator);
        String f1 = normalizeFormulaForWrite(formula1);
        String f2 = normalizeFormulaForWrite(formula2);
        if (ruleNode.path("bizInfo").path("showTime").asBoolean(false)) {
            return helper.createTimeConstraint(op, f1, f2);
        }
        return helper.createDateConstraint(op, f1, f2, "yyyy-MM-dd");
    }

    private static DataValidationConstraint createListConstraint(DataValidationHelper helper,
                                                                 String formula1,
                                                                 String formula2,
                                                                 boolean showDropDown) {
        String f1 = normalizeFormulaForWrite(formula1);
        if (f1 == null || f1.isEmpty()) {
            return helper.createExplicitListConstraint(new String[0]);
        }
        if (!f1.startsWith("=") && !f1.contains(":")) {
            String[] values = f1.split(",", -1);
            DataValidationConstraint constraint = helper.createExplicitListConstraint(values);
            if (!showDropDown && constraint != null) {
                constraint.setFormula2(formula2);
            }
            return constraint;
        }
        return helper.createFormulaListConstraint(f1.startsWith("=") ? f1 : "=" + f1);
    }

    private static String mapType(int type) {
        switch (type) {
            case DataValidationConstraint.ValidationType.INTEGER:
                return "whole";
            case DataValidationConstraint.ValidationType.DECIMAL:
                return "decimal";
            case DataValidationConstraint.ValidationType.LIST:
                return "list";
            case DataValidationConstraint.ValidationType.DATE:
                return "date";
            case DataValidationConstraint.ValidationType.TIME:
                return "date";
            case DataValidationConstraint.ValidationType.TEXT_LENGTH:
                return "textLength";
            case DataValidationConstraint.ValidationType.FORMULA:
                return "custom";
            case DataValidationConstraint.ValidationType.ANY:
            default:
                return "any";
        }
    }

    private static String mapOperator(int operator) {
        if (operator == DataValidationConstraint.OperatorType.BETWEEN) {
            return "between";
        }
        if (operator == DataValidationConstraint.OperatorType.EQUAL) {
            return "equal";
        }
        if (operator == DataValidationConstraint.OperatorType.NOT_BETWEEN) {
            return "notBetween";
        }
        if (operator == DataValidationConstraint.OperatorType.NOT_EQUAL) {
            return "notEqual";
        }
        if (operator == DataValidationConstraint.OperatorType.GREATER_THAN) {
            return "greaterThan";
        }
        if (operator == DataValidationConstraint.OperatorType.GREATER_OR_EQUAL) {
            return "greaterThanOrEqual";
        }
        if (operator == DataValidationConstraint.OperatorType.LESS_THAN) {
            return "lessThan";
        }
        if (operator == DataValidationConstraint.OperatorType.LESS_OR_EQUAL) {
            return "lessThanOrEqual";
        }
        return null;
    }

    private static int mapOperatorToPoi(String operator) {
        if (operator == null) {
            return DataValidationConstraint.OperatorType.IGNORED;
        }
        switch (operator) {
            case "between":
                return DataValidationConstraint.OperatorType.BETWEEN;
            case "equal":
                return DataValidationConstraint.OperatorType.EQUAL;
            case "notBetween":
                return DataValidationConstraint.OperatorType.NOT_BETWEEN;
            case "notEqual":
                return DataValidationConstraint.OperatorType.NOT_EQUAL;
            case "greaterThan":
                return DataValidationConstraint.OperatorType.GREATER_THAN;
            case "greaterThanOrEqual":
                return DataValidationConstraint.OperatorType.GREATER_OR_EQUAL;
            case "lessThan":
                return DataValidationConstraint.OperatorType.LESS_THAN;
            case "lessThanOrEqual":
                return DataValidationConstraint.OperatorType.LESS_OR_EQUAL;
            default:
                return DataValidationConstraint.OperatorType.IGNORED;
        }
    }

    private static String mapErrorStyle(int style) {
        switch (style) {
            case DataValidation.ErrorStyle.WARNING:
                return "warning";
            case DataValidation.ErrorStyle.INFO:
                return "info";
            case DataValidation.ErrorStyle.STOP:
            default:
                return "stop";
        }
    }

    private static int mapErrorStyleFromJson(JsonNode node) {
        if (node == null || node.isMissingNode() || node.isNull()) {
            return DataValidation.ErrorStyle.STOP;
        }
        if (node.isNumber()) {
            return node.asInt();
        }
        switch (node.asText()) {
            case "warning":
                return DataValidation.ErrorStyle.WARNING;
            case "info":
                return DataValidation.ErrorStyle.INFO;
            case "stop":
            default:
                return DataValidation.ErrorStyle.STOP;
        }
    }

    private static String extractFormula1(DataValidationConstraint constraint) {
        if (constraint == null) {
            return null;
        }
        if (constraint.getValidationType() == DataValidationConstraint.ValidationType.LIST) {
            String[] explicit = constraint.getExplicitListValues();
            if (explicit != null && explicit.length > 0) {
                return joinListValues(explicit);
            }
            String formula = constraint.getFormula1();
            if (formula == null || formula.isEmpty()) {
                return null;
            }
            return formula.startsWith("=") ? formula : "=" + formula;
        }
        return normalizeFormulaForRead(constraint, constraint.getFormula1());
    }

    private static String normalizeFormulaForRead(DataValidationConstraint constraint, String formula) {
        if (formula == null || formula.isEmpty()) {
            return null;
        }
        return formula;
    }

    private static String normalizeFormulaForWrite(String formula) {
        if (formula == null) {
            return null;
        }
        String trimmed = formula.trim();
        if (trimmed.isEmpty()) {
            return null;
        }
        return trimmed;
    }

    /**
     * 将显式列表值拼接为逗号分隔字符串。
     * Join explicit list values into a comma-separated string.
     * <p>
     * 已知限制 / Known limitations: 值中包含逗号或引号时无法正确转义，会导致解析错误。
     * Values containing commas or quotes cannot be escaped and will cause parsing errors.
     */
    private static String joinListValues(String[] values) {
        if (values == null || values.length == 0) {
            return null;
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < values.length; i++) {
            if (i > 0) {
                sb.append(',');
            }
            sb.append(values[i] == null ? "" : values[i]);
        }
        return sb.toString();
    }

    private static ArrayNode rangesToJson(CellRangeAddressList list, ObjectMapper mapper) {
        ArrayNode out = mapper.createArrayNode();
        if (list == null) {
            return out;
        }
        for (CellRangeAddress r : list.getCellRangeAddresses()) {
            if (r == null) {
                continue;
            }
            ObjectNode range = mapper.createObjectNode();
            range.put("startRow", r.getFirstRow());
            range.put("startColumn", r.getFirstColumn());
            range.put("endRow", r.getLastRow());
            range.put("endColumn", r.getLastColumn());
            range.put("startAbsoluteRefType", 0);
            range.put("endAbsoluteRefType", 0);
            range.put("rangeType", 0);
            out.add(range);
        }
        return out;
    }

    private static CellRangeAddressList jsonToRanges(JsonNode rangesNode) {
        CellRangeAddressList list = new CellRangeAddressList();
        if (rangesNode == null || !rangesNode.isArray()) {
            return list;
        }
        for (JsonNode rangeNode : rangesNode) {
            int startRow = rangeNode.path("startRow").asInt(-1);
            int startColumn = rangeNode.path("startColumn").asInt(-1);
            int endRow = rangeNode.path("endRow").asInt(-1);
            int endColumn = rangeNode.path("endColumn").asInt(-1);
            if (startRow < 0 || startColumn < 0 || endRow < 0 || endColumn < 0) {
                continue;
            }
            list.addCellRangeAddress(startRow, startColumn, endRow, endColumn);
        }
        return list;
    }

    private static Iterable<Map.Entry<String, JsonNode>> iterable(final java.util.Iterator<Map.Entry<String, JsonNode>> it) {
        return () -> it;
    }

    /**
     * Build a sheet-level resource object for {@code resources}.
     * Build the plugin resource payload keyed by sheet id.
     */
    public static ObjectNode buildResourceBySheetId(XSSFSheet sheet, ObjectMapper mapper) {
        ObjectNode out = mapper.createObjectNode();
        ObjectNode rules = readSheetDataValidations(sheet, mapper);
        if (rules.size() == 0) {
            return out;
        }
        out.set(String.valueOf(sheet.getWorkbook().getSheetIndex(sheet)), rules);
        return out;
    }
}
