package io.github.autoffice.univer.model;

import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class CompositePojoJsonTest {
    private final ObjectMapper m = JsonMapper.get();

    @Test
    void should_deserialize_cell_data_with_numeric_keys() throws Exception {
        String json = "{\"cellData\":{\"0\":{\"0\":{\"v\":\"A1\"},\"1\":{\"v\":1}}}}";
        IWorksheetData ws = m.readValue(json, IWorksheetData.class);
        assertThat(ws.getCellData().get(0).get(0).getV()).isEqualTo("A1");
        assertThat(ws.getCellData().get(0).get(1).getV()).isEqualTo(1);
    }

    @Test
    void should_roundtrip_minimal_workbook() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb1").setName("demo")
                .setAppVersion("0.10.2").setLocale("enUS");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("Sheet1"));
        wb.setSheetOrder(Collections.singletonList("s1"));
        String json = m.writeValueAsString(wb);
        IWorkbookData back = m.readValue(json, IWorkbookData.class);
        assertThat(back.getId()).isEqualTo("wb1");
        assertThat(back.getSheets().get("s1").getName()).isEqualTo("Sheet1");
    }

    @Test
    void should_roundtrip_style_data() throws Exception {
        IStyleData s = new IStyleData().setFf("Arial").setFs(12).setBl(BooleanNumber.TRUE)
                .setCl(new IColorStyle().setRgb("#ff0000"))
                .setN(new INumfmtLocal().setPattern("yyyy-mm-dd"));
        String json = m.writeValueAsString(s);
        IStyleData back = m.readValue(json, IStyleData.class);
        assertThat(back.getFf()).isEqualTo("Arial");
        assertThat(back.getN().getPattern()).isEqualTo("yyyy-mm-dd");
    }

    @Test
    void should_roundtrip_cell_with_formula_and_custom() throws Exception {
        Map<String, Object> custom = new LinkedHashMap<>();
        custom.put("key", "value");
        ICellData cell = new ICellData().setV(42.0).setT(CellValueType.NUMBER)
                .setF("SUM(A1:B1)").setSi("sid1").setCustom(custom);
        String json = m.writeValueAsString(cell);
        ICellData back = m.readValue(json, ICellData.class);
        assertThat(back.getF()).isEqualTo("SUM(A1:B1)");
        assertThat(back.getCustom()).containsEntry("key", "value");
    }
}
