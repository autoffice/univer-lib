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
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

/**
 * 工作表级自动筛选转换器：在 POI XSSFSheet 的 autoFilter 与 Univer
 * {@code SHEET_FILTER_PLUGIN} 插件资源之间做 best-effort 双向映射。
 *
 * <p>Univer 侧 per-sheet payload 结构：
 * <pre>
 * {
 *   "ref": "A1:D10",
 *   "startRow": 0,
 *   "endRow": 9,
 *   "startColumn": 0,
 *   "endColumn": 3,
 *   "rawXml": "&lt;autoFilter xmlns=\"...\" ref=\"A1:D10\"&gt;...&lt;/autoFilter&gt;"
 * }
 * </pre>
 *
 * <p>仅 {@code ref} 是权威字段；{@code rawXml} 用于无损保存列级 filter/sort 细节，
 * 因为 Univer 的列级筛选 schema 未公开，且 poi-ooxml-lite 裁剪了 CTFilterColumn 相关类。
 *
 * <p>Best-effort sheet-level auto filter bridge between POI XSSFSheet autoFilter
 * and Univer's {@code SHEET_FILTER_PLUGIN} plugin resource. Column-level criteria
 * are not modelled explicitly; instead, the raw OOXML is preserved via
 * {@code rawXml} to guarantee round-trip fidelity.
 */
public final class FilterConverter {

    /** Univer 自动筛选插件名 / Univer filter plugin resource name. */
    public static final String PLUGIN_NAME = "SHEET_FILTER_PLUGIN";

    private FilterConverter() {
    }

    // ============================================================
    // 读路径 / Read path: POI -> Univer payload JSON
    // ============================================================

    /**
     * 读取 sheet 的自动筛选区域并转为 per-sheet 资源；无 autoFilter 时返回空对象。
     * Read a sheet's auto filter and convert it to a per-sheet resource payload; returns
     * an empty object if the sheet has no auto filter.
     */
    public static ObjectNode readSheetFilter(XSSFSheet sheet, ObjectMapper mapper) {
        ObjectNode out = mapper.createObjectNode();
        if (sheet == null) {
            return out;
        }
        CTWorksheet cw = sheet.getCTWorksheet();
        if (cw == null || !cw.isSetAutoFilter()) {
            return out;
        }
        CTAutoFilter af = cw.getAutoFilter();
        if (af == null) {
            return out;
        }
        String ref = af.getRef();
        if (ref == null || ref.isEmpty()) {
            return out;
        }
        out.put("ref", ref);
        CellRangeAddress range = parseRange(ref);
        if (range != null) {
            out.put("startRow", range.getFirstRow());
            out.put("endRow", range.getLastRow());
            out.put("startColumn", range.getFirstColumn());
            out.put("endColumn", range.getLastColumn());
        }
        // 保留原 OOXML XML 以便写回时无损还原列级 filter / sortState。
        // Preserve the original OOXML so writing back can restore column-level detail losslessly.
        String xml = af.xmlText();
        if (xml != null && !xml.isEmpty()) {
            out.put("rawXml", xml);
        }
        return out;
    }

    // ============================================================
    // 写路径 / Write path: Univer payload JSON -> POI autoFilter
    // ============================================================

    /**
     * 把 Univer 自动筛选 payload 写回 sheet。payload 至少需包含 {@code ref} 或
     * {@code startRow/endRow/startColumn/endColumn}；存在 {@code rawXml} 时优先使用原 XML。
     * Apply a Univer auto filter payload back onto a POI sheet. Requires either
     * {@code ref} or the numeric range fields; {@code rawXml} is preferred when present.
     */
    public static void writeSheetFilter(XSSFSheet sheet, JsonNode payload) {
        if (sheet == null || payload == null || !payload.isObject() || payload.size() == 0) {
            return;
        }
        CellRangeAddress range = resolveRange(payload);
        if (range == null) {
            return;
        }
        // 先用 POI API 建立 autoFilter + _xlnm._FilterDatabase defined name；
        // 随后如有 rawXml，则覆盖 autoFilter 节点以还原列级 filter/sort 细节。
        // Establish the autoFilter node + _xlnm._FilterDatabase defined name via POI first,
        // then overlay rawXml to restore column-level detail when present.
        sheet.setAutoFilter(range);
        String rawXml = payload.path("rawXml").asText(null);
        if (rawXml == null || rawXml.isEmpty()) {
            return;
        }
        CTWorksheet cw = sheet.getCTWorksheet();
        if (cw == null) {
            return;
        }
        try {
            CTAutoFilter parsed = CTAutoFilter.Factory.parse(rawXml);
            cw.setAutoFilter(parsed);
        } catch (Exception ignored) {
            // rawXml 坏了就保留上一步已经建好的 autoFilter，range-only 仍然生效。
            // Ignore malformed rawXml; keep the range-only autoFilter set above.
        }
    }

    private static CellRangeAddress resolveRange(JsonNode payload) {
        String ref = payload.path("ref").asText(null);
        if (ref != null && !ref.isEmpty()) {
            CellRangeAddress parsed = parseRange(ref);
            if (parsed != null) {
                return parsed;
            }
        }
        if (payload.hasNonNull("startRow") && payload.hasNonNull("endRow")
                && payload.hasNonNull("startColumn") && payload.hasNonNull("endColumn")) {
            try {
                return new CellRangeAddress(
                        payload.get("startRow").asInt(),
                        payload.get("endRow").asInt(),
                        payload.get("startColumn").asInt(),
                        payload.get("endColumn").asInt());
            } catch (RuntimeException ignored) {
                return null;
            }
        }
        return null;
    }

    private static CellRangeAddress parseRange(String ref) {
        try {
            return CellRangeAddress.valueOf(ref);
        } catch (IllegalArgumentException | IllegalStateException e) {
            return null;
        }
    }
}
