/*
 * Copyright © 2026 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.ColorUtils;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionFilterType;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.DataBarFormatting;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFDataBarFormatting;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;

import java.util.UUID;

/**
 * 条件格式转换器：在 POI 的 XSSFSheet 条件格式与 Univer 的
 * {@code SHEET_CONDITIONAL_FORMATTING_PLUGIN} 插件数据之间双向转换。
 *
 * <p>Univer 侧每条规则的结构：
 * <pre>
 * {
 *   "cfId": "xxxx",
 *   "ranges": [{"startRow":..,"startColumn":..,"endRow":..,"endColumn":..,
 *                "startAbsoluteRefType":0,"endAbsoluteRefType":0,"rangeType":0}],
 *   "rule": {...},
 *   "stopIfTrue": false
 * }
 * </pre>
 *
 * <p>覆盖 Univer 四大规则类型：highlightCell / colorScale / dataBar / iconSet。
 * 其它类型 (FILTER 等) 作 best-effort 处理，无法表达时跳过。
 *
 * <p>Conditional formatting bridge between POI's XSSFSheet and Univer's
 * {@code SHEET_CONDITIONAL_FORMATTING_PLUGIN} plugin payload. Handles
 * highlightCell, colorScale, dataBar, iconSet; other types degrade gracefully.
 */
public final class ConditionalFormattingConverter {

    /** Univer iconSet 默认尺寸上限，用于表达 "小于等于任意值"。 */
    private static final double ICON_SENTINEL_MAX = 9007199254740991d;

    private ConditionalFormattingConverter() {}

    // ============================================================
    // 读路径 / Read path: POI -> Univer JSON array
    // ============================================================

    /**
     * 读取 sheet 的全部条件格式，转为 Univer 规则数组。
     * Read all conditional formattings from a sheet, converting to a Univer rule JSON array.
     *
     * @return 规则数组；若 sheet 无条件格式则返回空数组，永不返回 null。
     */
    public static ArrayNode readSheetCf(XSSFSheet sheet, ObjectMapper mapper) {
        ArrayNode out = mapper.createArrayNode();
        if (sheet == null) {
            return out;
        }
        XSSFSheetConditionalFormatting cfGroup = sheet.getSheetConditionalFormatting();
        int cfCount = cfGroup.getNumConditionalFormattings();
        for (int i = 0; i < cfCount; i++) {
            ConditionalFormatting cf = cfGroup.getConditionalFormattingAt(i);
            CellRangeAddress[] ranges = cf.getFormattingRanges();
            int ruleCount = cf.getNumberOfRules();
            for (int r = 0; r < ruleCount; r++) {
                ConditionalFormattingRule rule = cf.getRule(r);
                ObjectNode univerRule = convertRuleToUniver(rule, ranges, mapper);
                if (univerRule != null) {
                    out.add(univerRule);
                }
            }
        }
        return out;
    }

    private static ObjectNode convertRuleToUniver(ConditionalFormattingRule rule,
                                                  CellRangeAddress[] ranges,
                                                  ObjectMapper mapper) {
        ConditionType type = rule.getConditionType();
        if (type == null) {
            return null;
        }
        ObjectNode univerRule = mapper.createObjectNode();
        univerRule.put("cfId", shortId());
        univerRule.set("ranges", rangesToJson(ranges, mapper));

        ObjectNode ruleBody;
        if (ConditionType.COLOR_SCALE.equals(type)) {
            ruleBody = buildColorScaleRule(rule, mapper);
        } else if (ConditionType.DATA_BAR.equals(type)) {
            ruleBody = buildDataBarRule(rule, mapper);
        } else if (ConditionType.ICON_SET.equals(type)) {
            ruleBody = buildIconSetRule(rule, mapper);
        } else if (ConditionType.FORMULA.equals(type)) {
            ruleBody = buildHighlightFormulaRule(rule, mapper);
        } else if (ConditionType.CELL_VALUE_IS.equals(type)) {
            ruleBody = buildHighlightNumberRule(rule, mapper);
        } else if (ConditionType.FILTER.equals(type)) {
            ruleBody = buildFilterRule(rule, mapper);
        } else {
            return null;
        }
        if (ruleBody == null) {
            return null;
        }
        univerRule.set("rule", ruleBody);
        univerRule.put("stopIfTrue", rule.getStopIfTrue());
        return univerRule;
    }

    private static ObjectNode buildColorScaleRule(ConditionalFormattingRule rule, ObjectMapper mapper) {
        ColorScaleFormatting cs = rule.getColorScaleFormatting();
        if (cs == null) {
            return null;
        }
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "colorScale");
        ArrayNode config = mapper.createArrayNode();
        org.apache.poi.ss.usermodel.Color[] colors = cs.getColors();
        ConditionalFormattingThreshold[] thresholds = cs.getThresholds();
        int n = Math.min(colors == null ? 0 : colors.length, thresholds == null ? 0 : thresholds.length);
        for (int i = 0; i < n; i++) {
            ObjectNode stop = mapper.createObjectNode();
            String color = extractHex(colors[i]);
            if (color != null) {
                stop.put("color", color);
            }
            stop.set("value", thresholdToValue(thresholds[i], mapper));
            stop.put("index", i);
            config.add(stop);
        }
        body.set("config", config);
        return body;
    }

    private static ObjectNode buildDataBarRule(ConditionalFormattingRule rule, ObjectMapper mapper) {
        DataBarFormatting db = rule.getDataBarFormatting();
        if (db == null) {
            return null;
        }
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "dataBar");
        ObjectNode config = mapper.createObjectNode();
        config.set("min", thresholdToValue(db.getMinThreshold(), mapper));
        config.set("max", thresholdToValue(db.getMaxThreshold(), mapper));
        config.put("isGradient", true);
        String posColor = null;
        if (db instanceof XSSFDataBarFormatting) {
            XSSFColor xc = ((XSSFDataBarFormatting) db).getColor();
            posColor = ColorUtils.argbToRgbHex(xc == null ? null : xc.getARGB());
        }
        if (posColor == null) {
            posColor = "#638EC6";
        }
        config.put("positiveColor", posColor);
        config.put("nativeColor", posColor);
        body.set("config", config);
        return body;
    }

    private static ObjectNode buildIconSetRule(ConditionalFormattingRule rule, ObjectMapper mapper) {
        IconMultiStateFormatting ic = rule.getMultiStateFormatting();
        if (ic == null) {
            return null;
        }
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "iconSet");
        body.put("isShowValue", !ic.isIconOnly());
        String iconType = ic.getIconSet() != null ? ic.getIconSet().name : "3Arrows";
        ArrayNode config = mapper.createArrayNode();
        ConditionalFormattingThreshold[] thresholds = ic.getThresholds();
        int n = thresholds == null ? 0 : thresholds.length;
        for (int i = 0; i < n; i++) {
            ObjectNode stop = mapper.createObjectNode();
            // Univer 用 operator 描述区间上下界；POI 多阶阈值语义是 "下一阶之前的下限"
            String op = (i == n - 1) ? "lessThanOrEqual" : "greaterThan";
            stop.put("operator", op);
            stop.set("value", thresholdToValue(thresholds[i], mapper));
            stop.put("iconType", iconType);
            stop.put("iconId", String.valueOf(i));
            config.add(stop);
        }
        body.set("config", config);
        return body;
    }

    private static ObjectNode buildHighlightFormulaRule(ConditionalFormattingRule rule, ObjectMapper mapper) {
        ObjectNode body = mapper.createObjectNode();
        body.set("style", extractStyle(rule, mapper));
        String f = rule.getFormula1();
        body.put("value", f == null ? "" : ("=" + stripLeadingEq(f)));
        body.put("type", "highlightCell");
        body.put("subType", "formula");
        return body;
    }

    private static ObjectNode buildHighlightNumberRule(ConditionalFormattingRule rule, ObjectMapper mapper) {
        ObjectNode body = mapper.createObjectNode();
        body.set("style", extractStyle(rule, mapper));
        body.put("type", "highlightCell");
        body.put("subType", "number");
        body.put("operator", mapComparisonOperator(rule.getComparisonOperation()));
        String f1 = rule.getFormula1();
        String f2 = rule.getFormula2();
        if (f1 != null) {
            body.put("value", f1);
        }
        if (f2 != null) {
            body.put("value2", f2);
        }
        return body;
    }

    private static ObjectNode buildFilterRule(ConditionalFormattingRule rule, ObjectMapper mapper) {
        ConditionFilterType ft = rule.getConditionFilterType();
        if (ft == null) {
            return null;
        }
        ObjectNode body = mapper.createObjectNode();
        body.set("style", extractStyle(rule, mapper));
        body.put("type", "highlightCell");
        switch (ft) {
            case CONTAINS_TEXT:
                body.put("subType", "text");
                body.put("operator", "containsText");
                if (rule.getText() != null) {
                    body.put("value", rule.getText());
                }
                return body;
            case NOT_CONTAINS_TEXT:
                body.put("subType", "text");
                body.put("operator", "notContainsText");
                if (rule.getText() != null) {
                    body.put("value", rule.getText());
                }
                return body;
            case BEGINS_WITH:
                body.put("subType", "text");
                body.put("operator", "beginsWith");
                if (rule.getText() != null) {
                    body.put("value", rule.getText());
                }
                return body;
            case ENDS_WITH:
                body.put("subType", "text");
                body.put("operator", "endsWith");
                if (rule.getText() != null) {
                    body.put("value", rule.getText());
                }
                return body;
            case DUPLICATE_VALUES:
                body.put("subType", "duplicateValues");
                return body;
            case UNIQUE_VALUES:
                body.put("subType", "uniqueValues");
                return body;
            default:
                return null;
        }
    }

    // ============================================================
    // 辅助：读路径工具 / Read-path helpers
    // ============================================================

    private static ArrayNode rangesToJson(CellRangeAddress[] ranges, ObjectMapper mapper) {
        ArrayNode arr = mapper.createArrayNode();
        if (ranges == null) {
            return arr;
        }
        for (CellRangeAddress r : ranges) {
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
            arr.add(range);
        }
        return arr;
    }

    private static ObjectNode thresholdToValue(ConditionalFormattingThreshold t, ObjectMapper mapper) {
        ObjectNode v = mapper.createObjectNode();
        if (t == null) {
            v.put("type", "num");
            v.put("value", 0);
            return v;
        }
        ConditionalFormattingThreshold.RangeType rt = t.getRangeType();
        if (rt == null) {
            v.put("type", "num");
        } else {
            v.put("type", mapRangeType(rt));
        }
        if (t.getValue() != null) {
            v.put("value", t.getValue());
        } else if (t.getFormula() != null) {
            v.put("value", t.getFormula());
        } else {
            v.put("value", 0);
        }
        return v;
    }

    private static String mapRangeType(ConditionalFormattingThreshold.RangeType rt) {
        switch (rt) {
            case MIN: return "min";
            case MAX: return "max";
            case PERCENT: return "percent";
            case PERCENTILE: return "percentile";
            case FORMULA: return "formula";
            case NUMBER:
            default: return "num";
        }
    }

    private static String extractHex(org.apache.poi.ss.usermodel.Color color) {
        if (color instanceof XSSFColor) {
            return ColorUtils.argbToRgbHex(((XSSFColor) color).getARGB());
        }
        if (color instanceof ExtendedColor) {
            ExtendedColor ec = (ExtendedColor) color;
            try {
                return ColorUtils.argbToRgbHex(ec.getARGB());
            } catch (Exception ignored) {
                return null;
            }
        }
        return null;
    }

    private static ObjectNode extractStyle(ConditionalFormattingRule rule, ObjectMapper mapper) {
        ObjectNode style = mapper.createObjectNode();
        FontFormatting ff = rule.getFontFormatting();
        if (ff != null) {
            String fc = extractHex(ff.getFontColor());
            if (fc != null) {
                ObjectNode cl = mapper.createObjectNode();
                cl.put("rgb", fc);
                style.set("cl", cl);
            }
            if (ff.isBold()) {
                style.put("bl", 1);
            }
            if (ff.isItalic()) {
                style.put("it", 1);
            }
            if (ff.isStruckout()) {
                ObjectNode st = mapper.createObjectNode();
                st.put("s", 1);
                style.set("st", st);
            }
        }
        PatternFormatting pf = rule.getPatternFormatting();
        if (pf != null) {
            String bg = extractHex(pf.getFillBackgroundColorColor());
            if (bg == null) {
                bg = extractHex(pf.getFillForegroundColorColor());
            }
            if (bg != null) {
                ObjectNode bgNode = mapper.createObjectNode();
                bgNode.put("rgb", bg);
                style.set("bg", bgNode);
            }
        }
        return style;
    }

    private static String mapComparisonOperator(byte op) {
        switch (op) {
            case ComparisonOperator.EQUAL:       return "equal";
            case ComparisonOperator.NOT_EQUAL:   return "notEqual";
            case ComparisonOperator.GT:          return "greaterThan";
            case ComparisonOperator.LT:          return "lessThan";
            case ComparisonOperator.GE:          return "greaterThanOrEqual";
            case ComparisonOperator.LE:          return "lessThanOrEqual";
            case ComparisonOperator.BETWEEN:     return "between";
            case ComparisonOperator.NOT_BETWEEN: return "notBetween";
            default: return "equal";
        }
    }

    private static String stripLeadingEq(String s) {
        return (s != null && s.startsWith("=")) ? s.substring(1) : s;
    }

    private static String shortId() {
        return UUID.randomUUID().toString().replace("-", "").substring(0, 8);
    }

    // ============================================================
    // 写路径 / Write path: Univer JSON array -> POI
    // ============================================================

    /**
     * 将 Univer 规则数组写入 sheet 的条件格式。
     * Apply a Univer rule JSON array to the sheet as conditional formatting.
     */
    public static void writeSheetCf(XSSFSheet sheet, JsonNode rulesArray) {
        if (sheet == null || rulesArray == null || !rulesArray.isArray() || rulesArray.size() == 0) {
            return;
        }
        XSSFSheetConditionalFormatting cfGroup = sheet.getSheetConditionalFormatting();
        for (JsonNode ruleNode : rulesArray) {
            if (ruleNode == null || !ruleNode.isObject()) {
                continue;
            }
            CellRangeAddress[] ranges = parseRanges(ruleNode.get("ranges"));
            if (ranges == null || ranges.length == 0) {
                continue;
            }
            JsonNode body = ruleNode.get("rule");
            if (body == null || !body.isObject()) {
                continue;
            }
            boolean stopIfTrue = ruleNode.path("stopIfTrue").asBoolean(false);
            String type = body.path("type").asText("");
            XSSFConditionalFormattingRule created = null;
            switch (type) {
                case "highlightCell":
                    created = buildHighlightCell(cfGroup, body);
                    break;
                case "colorScale":
                    created = buildColorScale(cfGroup, body);
                    break;
                case "dataBar":
                    created = buildDataBar(cfGroup, body);
                    break;
                case "iconSet":
                    created = buildIconSet(cfGroup, body);
                    break;
                default:
                    continue;
            }
            if (created == null) {
                continue;
            }
            if (stopIfTrue) {
                // XSSFConditionalFormattingRule 的 getCTCfRule() 为 package-private，
                // 通过反射读取底层 CTCfRule 并设 stopIfTrue；失败则忽略（best-effort）。
                try {
                    java.lang.reflect.Method m = created.getClass().getDeclaredMethod("getCTCfRule");
                    m.setAccessible(true);
                    Object ct = m.invoke(created);
                    ct.getClass().getMethod("setStopIfTrue", boolean.class).invoke(ct, true);
                } catch (Exception ignored) {
                    // ignore
                }
            }
            cfGroup.addConditionalFormatting(ranges, created);
        }
    }

    private static CellRangeAddress[] parseRanges(JsonNode arr) {
        if (arr == null || !arr.isArray() || arr.size() == 0) {
            return null;
        }
        CellRangeAddress[] out = new CellRangeAddress[arr.size()];
        for (int i = 0; i < arr.size(); i++) {
            JsonNode n = arr.get(i);
            int sr = n.path("startRow").asInt(0);
            int sc = n.path("startColumn").asInt(0);
            int er = n.path("endRow").asInt(sr);
            int ec = n.path("endColumn").asInt(sc);
            out[i] = new CellRangeAddress(sr, er, sc, ec);
        }
        return out;
    }

    private static XSSFConditionalFormattingRule buildHighlightCell(XSSFSheetConditionalFormatting cfGroup,
                                                                    JsonNode body) {
        String subType = body.path("subType").asText("formula");
        XSSFConditionalFormattingRule rule;
        if ("formula".equals(subType)) {
            String formula = stripLeadingEq(body.path("value").asText(""));
            if (formula == null || formula.isEmpty()) {
                return null;
            }
            rule = cfGroup.createConditionalFormattingRule(formula);
        } else if ("number".equals(subType)) {
            byte op = univerOperatorToPoi(body.path("operator").asText("equal"));
            String v1 = body.path("value").asText("");
            if (op == ComparisonOperator.BETWEEN || op == ComparisonOperator.NOT_BETWEEN) {
                String v2 = body.path("value2").asText("");
                rule = cfGroup.createConditionalFormattingRule(op, v1, v2);
            } else {
                rule = cfGroup.createConditionalFormattingRule(op, v1);
            }
        } else {
            // text/duplicate/unique: fallback - use a dummy formula so at least a CF stub is written.
            // POI's XSSF high-level API has limited support for these; we degrade to a FORMULA rule.
            String text = body.path("value").asText("");
            String safe = text.replace("\"", "\"\"");
            rule = cfGroup.createConditionalFormattingRule("ISTEXT(A1)*LEN(\"" + safe + "\")>0");
        }
        applyStyle(rule, body.path("style"));
        return rule;
    }

    private static XSSFConditionalFormattingRule buildColorScale(XSSFSheetConditionalFormatting cfGroup,
                                                                 JsonNode body) {
        XSSFConditionalFormattingRule rule = cfGroup.createConditionalFormattingColorScaleRule();
        org.apache.poi.xssf.usermodel.XSSFColorScaleFormatting cs = rule.getColorScaleFormatting();
        JsonNode config = body.path("config");
        if (!config.isArray() || config.size() == 0) {
            return rule;
        }
        int n = config.size();
        cs.setNumControlPoints(n);
        ConditionalFormattingThreshold[] thresholds = cs.getThresholds();
        org.apache.poi.ss.usermodel.Color[] colors = new org.apache.poi.ss.usermodel.Color[n];
        for (int i = 0; i < n; i++) {
            JsonNode stop = config.get(i);
            applyThreshold(thresholds[i], stop.path("value"), i == 0, i == n - 1);
            XSSFColor xc = newXssfColor(stop.path("color").asText(null));
            colors[i] = xc;
        }
        cs.setThresholds(thresholds);
        cs.setColors(colors);
        return rule;
    }

    private static XSSFConditionalFormattingRule buildDataBar(XSSFSheetConditionalFormatting cfGroup,
                                                              JsonNode body) {
        JsonNode config = body.path("config");
        String colorHex = config.path("positiveColor").asText(null);
        if (colorHex == null) {
            colorHex = config.path("nativeColor").asText("#638EC6");
        }
        XSSFColor xc = newXssfColor(colorHex);
        if (xc == null) {
            xc = newXssfColor("#638EC6");
        }
        XSSFConditionalFormattingRule rule = cfGroup.createConditionalFormattingRule(xc);
        DataBarFormatting db = rule.getDataBarFormatting();
        applyThreshold(db.getMinThreshold(), config.path("min"), true, false);
        applyThreshold(db.getMaxThreshold(), config.path("max"), false, true);
        return rule;
    }

    private static XSSFConditionalFormattingRule buildIconSet(XSSFSheetConditionalFormatting cfGroup,
                                                              JsonNode body) {
        JsonNode config = body.path("config");
        if (!config.isArray() || config.size() == 0) {
            return null;
        }
        String iconTypeName = config.get(0).path("iconType").asText("3Arrows");
        IconMultiStateFormatting.IconSet iconSet = mapIconSet(iconTypeName, config.size());
        XSSFConditionalFormattingRule rule = cfGroup.createConditionalFormattingRule(iconSet);
        IconMultiStateFormatting ic = rule.getMultiStateFormatting();
        boolean iconOnly = !body.path("isShowValue").asBoolean(true);
        ic.setIconOnly(iconOnly);
        ConditionalFormattingThreshold[] thresholds = ic.getThresholds();
        int n = Math.min(thresholds.length, config.size());
        for (int i = 0; i < n; i++) {
            JsonNode stop = config.get(i);
            applyThreshold(thresholds[i], stop.path("value"), i == 0, i == n - 1);
        }
        ic.setThresholds(thresholds);
        return rule;
    }

    private static IconMultiStateFormatting.IconSet mapIconSet(String name, int count) {
        if (name == null) {
            name = "";
        }
        try {
            IconMultiStateFormatting.IconSet match = IconMultiStateFormatting.IconSet.byName(name);
            if (match != null) {
                return match;
            }
        } catch (Exception ignored) {
            // fall through to count-based fallback
        }
        // fallbacks based on count
        if (count == 4) {
            return IconMultiStateFormatting.IconSet.GYR_4_ARROWS;
        }
        if (count == 5) {
            return IconMultiStateFormatting.IconSet.GYYYR_5_ARROWS;
        }
        return IconMultiStateFormatting.IconSet.GYR_3_ARROW;
    }

    private static void applyThreshold(ConditionalFormattingThreshold th,
                                       JsonNode valueNode,
                                       boolean isMinSlot,
                                       boolean isMaxSlot) {
        if (th == null) {
            return;
        }
        String t = valueNode.isObject() ? valueNode.path("type").asText("num") : "num";
        JsonNode v = valueNode.isObject() ? valueNode.path("value") : valueNode;
        ConditionalFormattingThreshold.RangeType rt;
        switch (t) {
            case "min":        rt = ConditionalFormattingThreshold.RangeType.MIN; break;
            case "max":        rt = ConditionalFormattingThreshold.RangeType.MAX; break;
            case "percent":    rt = ConditionalFormattingThreshold.RangeType.PERCENT; break;
            case "percentile": rt = ConditionalFormattingThreshold.RangeType.PERCENTILE; break;
            case "formula":    rt = ConditionalFormattingThreshold.RangeType.FORMULA; break;
            case "num":
            default:           rt = ConditionalFormattingThreshold.RangeType.NUMBER; break;
        }
        // POI 要求 NUMBER 只能用于非边界阈值；当首/末阈值是 NUMBER 且语义接近 min/max 时交由调用方，这里保留用户指定。
        th.setRangeType(rt);
        if (rt == ConditionalFormattingThreshold.RangeType.FORMULA) {
            th.setFormula(v.asText(""));
        } else if (rt != ConditionalFormattingThreshold.RangeType.MIN
                && rt != ConditionalFormattingThreshold.RangeType.MAX) {
            if (v.isNumber()) {
                th.setValue(v.asDouble());
            } else {
                try {
                    th.setValue(Double.parseDouble(v.asText("0")));
                } catch (NumberFormatException e) {
                    th.setValue(0d);
                }
            }
        }
        // MIN/MAX 阈值在 XML 中不需要 val 属性；POI 的 setValue(null) 在 val 本就缺失时会抛异常，
        // 所以直接不动原值（POI 新建阈值默认无 val）。
    }

    private static XSSFColor newXssfColor(String hex) {
        byte[] argb = ColorUtils.rgbHexToArgb(hex);
        if (argb == null) {
            return null;
        }
        return new XSSFColor(argb);
    }

    private static void applyStyle(XSSFConditionalFormattingRule rule, JsonNode styleNode) {
        if (styleNode == null || !styleNode.isObject() || styleNode.size() == 0) {
            return;
        }
        // 背景色
        JsonNode bg = styleNode.path("bg");
        String bgHex = bg.isObject() ? bg.path("rgb").asText(null) : null;
        if (bgHex != null) {
            PatternFormatting pf = rule.createPatternFormatting();
            XSSFColor xc = newXssfColor(bgHex);
            if (xc != null) {
                pf.setFillBackgroundColor(xc);
                pf.setFillForegroundColor(xc);
                pf.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            }
        }
        // 字体颜色 / 粗体 / 斜体 / 中划线
        JsonNode cl = styleNode.path("cl");
        String clHex = cl.isObject() ? cl.path("rgb").asText(null) : null;
        boolean bold = styleNode.path("bl").asInt(0) == 1;
        boolean italic = styleNode.path("it").asInt(0) == 1;
        JsonNode stNode = styleNode.path("st");
        boolean strike = stNode.isObject() && stNode.path("s").asInt(0) == 1;
        if (clHex != null || bold || italic || strike) {
            FontFormatting ff = rule.createFontFormatting();
            if (clHex != null) {
                XSSFColor xc = newXssfColor(clHex);
                if (xc != null) {
                    ff.setFontColor(xc);
                }
            }
            ff.setFontStyle(italic, bold);
            if (strike) {
                // FontFormatting#setFontStyle 不直接支持 strike；POI 会从 XML 推断，此处 best-effort。
            }
        }
    }

    private static byte univerOperatorToPoi(String op) {
        if (op == null) {
            return ComparisonOperator.EQUAL;
        }
        switch (op) {
            case "equal":              return ComparisonOperator.EQUAL;
            case "notEqual":           return ComparisonOperator.NOT_EQUAL;
            case "greaterThan":        return ComparisonOperator.GT;
            case "lessThan":           return ComparisonOperator.LT;
            case "greaterThanOrEqual": return ComparisonOperator.GE;
            case "lessThanOrEqual":    return ComparisonOperator.LE;
            case "between":            return ComparisonOperator.BETWEEN;
            case "notBetween":         return ComparisonOperator.NOT_BETWEEN;
            default: return ComparisonOperator.EQUAL;
        }
    }

    /** 暴露 iconSet 的 max 哨兵值，测试可据此断言 Univer 侧的 "fallback" 最大阈值。 */
    public static double iconSentinelMax() {
        return ICON_SENTINEL_MAX;
    }
}
