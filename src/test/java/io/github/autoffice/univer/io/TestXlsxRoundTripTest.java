package io.github.autoffice.univer.io;

import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 用真实 xlsx 资源做端到端转换与回写校验。
 * End-to-end conversion test using a real xlsx fixture.
 *
 * <p>流程：test.xlsx → IWorkbookData → JSON 文件 → 读回 IWorkbookData → output.xlsx →
 * 再读 output.xlsx → 与原 IWorkbookData 比较；用于回归保证 round-trip 不丢数据。
 */
class TestXlsxRoundTripTest {

    private static final String FIXTURE = "/test.xlsx";

    @Test
    void should_roundtrip_real_xlsx_via_json_and_back() throws Exception {
        // 产物目录：target/roundtrip-output/，方便人工检查；mvn clean 时自动清理。
        Path tmp = Paths.get("target/roundtrip-output");
        Files.createDirectories(tmp);

        // 1. xlsx → IWorkbookData
        IWorkbookData firstRead;
        try (InputStream in = getClass().getResourceAsStream(FIXTURE)) {
            assertThat(in).as("fixture %s 必须存在", FIXTURE).isNotNull();
            firstRead = UniverXlsx.read(in);
        }
        assertThat(firstRead).isNotNull();
        assertThat(firstRead.getSheets()).as("至少应有一个工作表").isNotEmpty();
        assertThat(firstRead.getSheetOrder()).isNotEmpty();

        // 2. IWorkbookData → JSON 文件
        ObjectMapper mapper = JsonMapper.get();
        Path jsonFile = tmp.resolve("test.workbook.json");
        mapper.writerWithDefaultPrettyPrinter().writeValue(jsonFile.toFile(), firstRead);
        assertThat(Files.size(jsonFile)).as("导出 JSON 不应为空").isGreaterThan(0L);

        // 3. JSON 文件 → IWorkbookData（验证 JSON 自身可被反序列化）
        IWorkbookData fromJson = mapper.readValue(jsonFile.toFile(), IWorkbookData.class);
        assertThat(fromJson)
                .usingRecursiveComparison()
                .ignoringFields("extras")
                .isEqualTo(firstRead);

        // 4. IWorkbookData → output.xlsx
        Path outXlsx = tmp.resolve("output.xlsx");
        UniverXlsx.write(fromJson, outXlsx);
        assertThat(Files.size(outXlsx)).as("output.xlsx 不应为空").isGreaterThan(0L);

        // 5. output.xlsx → IWorkbookData，再与第一次读出的快照对比
        IWorkbookData secondRead = UniverXlsx.read(outXlsx);

        // 5a. 关键工作簿级字段一致
        assertThat(secondRead.getSheetOrder())
                .as("sheetOrder 必须保留")
                .containsExactlyElementsOf(firstRead.getSheetOrder());
        assertThat(secondRead.getSheets().keySet())
                .as("sheet id 集合必须一致（不应出现 shadow 条目）")
                .containsExactlyInAnyOrderElementsOf(firstRead.getSheets().keySet());

        // 5b. 每个工作表的核心数据一致
        for (Map.Entry<String, IWorksheetData> e : firstRead.getSheets().entrySet()) {
            String sid = e.getKey();
            IWorksheetData expected = e.getValue();
            IWorksheetData actual = secondRead.getSheets().get(sid);
            assertThat(actual).as("sheet %s 必须存在", sid).isNotNull();

            assertThat(actual.getName()).as("sheet %s name", sid).isEqualTo(expected.getName());

            assertCellDataEqual(sid, expected, actual);

            java.util.List<io.github.autoffice.univer.model.IRange> expectedMerges = expected.getMergeData() == null
                    ? java.util.Collections.emptyList()
                    : expected.getMergeData();
            java.util.List<io.github.autoffice.univer.model.IRange> actualMerges = actual.getMergeData() == null
                    ? java.util.Collections.emptyList()
                    : actual.getMergeData();
            assertThat(actualMerges)
                    .as("sheet %s mergeData", sid)
                    .usingRecursiveFieldByFieldElementComparatorIgnoringFields("extras")
                    .containsExactlyInAnyOrderElementsOf(expectedMerges);
        }

        // 5c. 关键字段（appVersion、locale）应保留
        assertThat(secondRead.getLocale()).isEqualTo(firstRead.getLocale());
        if (firstRead.getAppVersion() != null) {
            assertThat(secondRead.getAppVersion()).isEqualTo(firstRead.getAppVersion());
        }

        // 5d. 专项验证：Cell sheet 的 A6（row=5, col=0）是一个 inline rich text，必须被识别为 p
        // Specific check: "Cell" sheet A6 should be recognized as inline rich text (p field populated).
        // test.xlsx 中 A6 内容为 "Inline Style Cell"，三段不同样式：
        //   "Inline" 红色粗体 / "Style" 黑色粗体带中划线 / "Cell" 黑色粗体
        IWorksheetData cellSheetFirst = findSheetByName(firstRead, "Cell");
        assertThat(cellSheetFirst).as("Cell 工作表必须存在").isNotNull();
        ICellData a6First = cellSheetFirst.getCellData().get(5).get(0);
        assertThat(a6First).as("Cell A6 必须存在").isNotNull();
        assertThat(a6First.getP())
                .as("Cell A6 应为 rich text（p 字段非空）")
                .isNotNull();
        assertThat(a6First.getP().getBody().getDataStream())
                .as("Cell A6 rich text 内容应为 'Inline Style Cell' + Univer 段落终止符")
                .isEqualTo("Inline Style Cell\r\n");
        assertThat(a6First.getP().getBody().getTextRuns())
                .as("Cell A6 应包含多个 text runs（红色粗体/黑色粗体中划线/黑色粗体）")
                .hasSizeGreaterThanOrEqualTo(2);

        // round-trip 后 Cell A6 同样应是 rich text
        IWorksheetData cellSheetSecond = findSheetByName(secondRead, "Cell");
        ICellData a6Second = cellSheetSecond.getCellData().get(5).get(0);
        assertThat(a6Second.getP())
                .as("round-trip 后 Cell A6 仍应为 rich text")
                .isNotNull();
        assertThat(a6Second.getP().getBody().getDataStream())
                .isEqualTo("Inline Style Cell\r\n");
    }

    /** 按 name 查找工作表（不区分 id 与 name 的差异）。 */
    private static IWorksheetData findSheetByName(IWorkbookData wb, String name) {
        for (IWorksheetData ws : wb.getSheets().values()) {
            if (name.equals(ws.getName())) {
                return ws;
            }
        }
        return null;
    }

    /** 比较两个工作表的 cellData，按值与类型逐格断言（忽略 extras）。 */
    private static void assertCellDataEqual(String sid, IWorksheetData expected, IWorksheetData actual) {
        Map<Integer, Map<Integer, ICellData>> exp = expected.getCellData();
        Map<Integer, Map<Integer, ICellData>> act = actual.getCellData();
        if (exp == null || exp.isEmpty()) {
            assertThat(act == null || act.isEmpty())
                    .as("sheet %s 期望空 cellData", sid)
                    .isTrue();
            return;
        }
        assertThat(act).as("sheet %s cellData 行集合", sid).containsKeys(exp.keySet().toArray(new Integer[0]));
        for (Map.Entry<Integer, Map<Integer, ICellData>> rowE : exp.entrySet()) {
            int row = rowE.getKey();
            Map<Integer, ICellData> expRow = rowE.getValue();
            Map<Integer, ICellData> actRow = act.get(row);
            assertThat(actRow).as("sheet %s row %d", sid, row).isNotNull();
            for (Map.Entry<Integer, ICellData> cellE : expRow.entrySet()) {
                int col = cellE.getKey();
                ICellData expCell = cellE.getValue();
                ICellData actCell = actRow.get(col);
                assertThat(actCell).as("sheet %s (%d,%d) 必须存在", sid, row, col).isNotNull();
                // 公式 cell 的 cached value 和类型可能在 round-trip 后变化（不做公式计算），跳过比较
                // Formula cells may have different cached values/types after round-trip, skip comparison
                if (expCell.getF() == null) {
                    assertThat(normalizeValue(actCell.getV()))
                            .as("sheet %s (%d,%d) 值", sid, row, col)
                            .isEqualTo(normalizeValue(expCell.getV()));
                    assertThat(actCell.getT())
                            .as("sheet %s (%d,%d) 类型", sid, row, col)
                            .isEqualTo(expCell.getT());
                }
                assertThat(actCell.getF())
                        .as("sheet %s (%d,%d) 公式", sid, row, col)
                        .isEqualTo(expCell.getF());
            }
        }
    }

    /** Jackson 反序列化数值 fall-back 到 Integer/Double，统一转 double 便于比较。 */
    private static Object normalizeValue(Object v) {
        if (v instanceof Number) {
            return ((Number) v).doubleValue();
        }
        return v;
    }
}
