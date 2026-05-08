package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.CellValueType;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IStyleData;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * 单元格数据转换器：在 ICellData 与 POI XSSFCell 之间双向映射。
 * Cell converter between ICellData and POI XSSFCell.
 * <p>
 * 不处理 shared formula (si) 分组 —— 由 SharedFormulaRegistry 负责。
 * Does NOT handle shared formula si grouping — that's handled by SharedFormulaRegistry.
 */
public final class CellConverter {

    private final StyleConverter styles;

    public CellConverter(StyleConverter styles) {
        this.styles = styles;
    }

    // ============================================================
    // 写路径 / Write path
    // ============================================================

    /**
     * 将 ICellData 写入 POI 单元格。
     * Write ICellData into POI cell.
     */
    public void writeCell(XSSFCell poiCell, ICellData src) {
        if (src == null) {
            return;
        }

        // 1. 公式 / Formula
        String formula = src.getF();
        if (formula != null) {
            // 去掉前导 '=' / strip leading '='
            if (formula.startsWith("=")) {
                formula = formula.substring(1);
            }
            poiCell.setCellFormula(formula);
            // 设置缓存值 / set cached value
            setCachedValue(poiCell, src);
        } else {
            // 2. 类型 + 值 / Type + value
            writeValue(poiCell, src);
        }

        // 3. 样式 / Style
        applyStyle(poiCell, src);
    }

    // ============================================================
    // 读路径 / Read path
    // ============================================================

    /**
     * 从 POI 单元格读取 ICellData。
     * Read ICellData from POI cell. Returns null for blank/empty cells.
     */
    public ICellData readCell(XSSFCell poiCell) {
        if (poiCell == null) {
            return null;
        }

        CellType cellType = poiCell.getCellType();

        // BLANK 或 _NONE 返回 null / return null for blank
        if (cellType == CellType.BLANK || cellType == CellType._NONE) {
            // 如果有非默认样式，仍然返回 / still return if non-default style
            if (hasNonDefaultStyle(poiCell)) {
                ICellData data = new ICellData();
                applyReadStyle(poiCell, data);
                return data;
            }
            return null;
        }

        ICellData data = new ICellData();

        switch (cellType) {
            case FORMULA:
                readFormulaCell(poiCell, data);
                break;
            case STRING:
                readStringCell(poiCell, data);
                break;
            case NUMERIC:
                data.setT(CellValueType.NUMBER);
                data.setV(poiCell.getNumericCellValue());
                break;
            case BOOLEAN:
                data.setT(CellValueType.BOOLEAN);
                data.setV(poiCell.getBooleanCellValue() ? 1 : 0);
                break;
            case ERROR:
                // 错误单元格作为字符串读取 / read error cell as string
                data.setT(CellValueType.STRING);
                data.setV(String.valueOf(poiCell.getErrorCellValue()));
                break;
            default:
                return null;
        }

        // 读取样式 / read style
        applyReadStyle(poiCell, data);

        return data;
    }

    // ============================================================
    // 写路径辅助方法 / Write-path helpers
    // ============================================================

    /**
     * 设置公式单元格的缓存值。
     */
    private void setCachedValue(XSSFCell poiCell, ICellData src) {
        Object v = src.getV();
        if (v == null) {
            return;
        }
        CellValueType t = src.getT();
        if (t == CellValueType.BOOLEAN) {
            poiCell.setCellValue(toBooleanValue(v));
        } else if (t == CellValueType.NUMBER || v instanceof Number) {
            poiCell.setCellValue(((Number) v).doubleValue());
        } else {
            poiCell.setCellValue(String.valueOf(v));
        }
    }

    /**
     * 写入非公式单元格的值。
     */
    private void writeValue(XSSFCell poiCell, ICellData src) {
        Object v = src.getV();
        if (v == null) {
            return;
        }
        CellValueType t = src.getT();

        if (t == CellValueType.FORCE_TEXT) {
            poiCell.setCellValue(String.valueOf(v));
        } else if (t == CellValueType.BOOLEAN) {
            poiCell.setCellValue(toBooleanValue(v));
        } else if (t == CellValueType.NUMBER || v instanceof Number) {
            poiCell.setCellValue(((Number) v).doubleValue());
        } else if (t == CellValueType.STRING || v instanceof String) {
            poiCell.setCellValue(String.valueOf(v));
        } else {
            // 自动检测 / auto-detect
            poiCell.setCellValue(String.valueOf(v));
        }
    }

    /**
     * 应用样式到单元格（写路径）。
     * FORCE_TEXT 需要 quotePrefix=true，通过 StyleConverter 的缓存变体避免产生大量样式对象。
     */
    private void applyStyle(XSSFCell poiCell, ICellData src) {
        boolean isForceText = src.getT() == CellValueType.FORCE_TEXT;
        Object styleObj = src.getS();

        if (styleObj instanceof IStyleData) {
            IStyleData styleData = (IStyleData) styleObj;
            if (isForceText) {
                poiCell.setCellStyle(styles.toPoiStyleWithQuotePrefix(styleData));
            } else {
                poiCell.setCellStyle(styles.toPoiStyle(styleData));
            }
        } else if (isForceText) {
            // 无内联样式但需要 quotePrefix；交给 StyleConverter 去重缓存
            // No inline style but still need quotePrefix; let StyleConverter cache a shared variant.
            poiCell.setCellStyle(styles.toPoiStyleWithQuotePrefix(null));
        }
        // 如果 styleObj 是 String（样式 id），由调用方处理，这里跳过
    }

    // ============================================================
    // 读路径辅助方法 / Read-path helpers
    // ============================================================

    /**
     * 读取公式单元格。
     */
    private void readFormulaCell(XSSFCell poiCell, ICellData data) {
        data.setF(poiCell.getCellFormula());

        // 读取缓存值类型 / read cached value type
        CellType cachedType = poiCell.getCachedFormulaResultType();
        switch (cachedType) {
            case STRING:
                data.setT(CellValueType.STRING);
                data.setV(poiCell.getStringCellValue());
                break;
            case NUMERIC:
                data.setT(CellValueType.NUMBER);
                data.setV(poiCell.getNumericCellValue());
                break;
            case BOOLEAN:
                data.setT(CellValueType.BOOLEAN);
                data.setV(poiCell.getBooleanCellValue() ? 1 : 0);
                break;
            case ERROR:
                data.setT(CellValueType.STRING);
                data.setV(String.valueOf(poiCell.getErrorCellValue()));
                break;
            default:
                break;
        }
    }

    /**
     * 读取字符串单元格，检测 FORCE_TEXT。
     */
    private void readStringCell(XSSFCell poiCell, ICellData data) {
        XSSFCellStyle cs = poiCell.getCellStyle();
        if (cs != null && cs.getQuotePrefixed()) {
            data.setT(CellValueType.FORCE_TEXT);
        } else {
            data.setT(CellValueType.STRING);
        }
        data.setV(poiCell.getStringCellValue());
    }

    /**
     * 读取样式并设置到 ICellData。
     */
    private void applyReadStyle(XSSFCell poiCell, ICellData data) {
        XSSFCellStyle cs = poiCell.getCellStyle();
        if (cs == null || cs.getIndex() == 0) {
            return;
        }
        IStyleData styleData = styles.fromPoiStyle(cs);
        String styleId = styles.styleIdOf(styleData);
        // 检查是否为空样式 / check if empty style
        String emptyId = styles.styleIdOf(new IStyleData());
        if (!styleId.equals(emptyId)) {
            data.setS(styleId);
        }
    }

    /**
     * 检查单元格是否有非默认样式。
     */
    private boolean hasNonDefaultStyle(XSSFCell poiCell) {
        XSSFCellStyle cs = poiCell.getCellStyle();
        return cs != null && cs.getIndex() != 0;
    }

    /**
     * 将各种类型的值转换为 boolean。
     */
    private boolean toBooleanValue(Object v) {
        if (v instanceof Boolean) {
            return (Boolean) v;
        }
        if (v instanceof Number) {
            return ((Number) v).intValue() != 0;
        }
        return Boolean.parseBoolean(String.valueOf(v));
    }
}
