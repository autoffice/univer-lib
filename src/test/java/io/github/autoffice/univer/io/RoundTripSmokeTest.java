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
package io.github.autoffice.univer.io;

import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.CellValueType;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IRange;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class RoundTripSmokeTest {

    @Test
    void should_roundtrip_minimal_workbook() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("w").setAppVersion("0.10.2").setLocale("zhCN");
        src.setSheetOrder(Collections.singletonList("s1"));
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("Sheet1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("A1").setT(CellValueType.STRING));
        row.put(1, new ICellData().setV(3.14).setT(CellValueType.NUMBER));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);

        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        UniverXlsx.write(src, buf);

        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(buf.toByteArray()));
        assertThat(back.getSheetOrder()).containsExactly("s1");
        assertThat(back.getSheets().get("s1").getCellData().get(0).get(0).getV()).isEqualTo("A1");
        assertThat(back.getAppVersion()).isEqualTo("0.10.2");
    }

    @Test
    void should_roundtrip_two_sheets_with_hidden() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("w2").setAppVersion("0.10.2").setLocale("enUS");
        src.setSheetOrder(Arrays.asList("s1", "s2"));
        IWorksheetData ws1 = new IWorksheetData().setId("s1").setName("Visible");
        IWorksheetData ws2 = new IWorksheetData().setId("s2").setName("Hidden").setHidden(BooleanNumber.TRUE);
        Map<Integer, ICellData> r1 = new LinkedHashMap<>();
        r1.put(0, new ICellData().setV("x").setT(CellValueType.STRING));
        ws1.getCellData().put(0, r1);
        src.getSheets().put("s1", ws1);
        src.getSheets().put("s2", ws2);

        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        UniverXlsx.write(src, buf);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(buf.toByteArray()));

        assertThat(back.getSheetOrder()).containsExactly("s1", "s2");
        assertThat(back.getSheets().get("s2").getHidden()).isEqualTo(BooleanNumber.TRUE);
    }

    @Test
    void should_roundtrip_with_styles_and_merges() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("w3").setAppVersion("0.10.2");
        src.setSheetOrder(Collections.singletonList("s1"));
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("Sheet1");
        IStyleData inline = new IStyleData().setFf("Arial").setFs(14).setBl(BooleanNumber.TRUE);
        Map<Integer, ICellData> r = new LinkedHashMap<>();
        r.put(0, new ICellData().setV("Bold").setT(CellValueType.STRING).setS(inline));
        ws.getCellData().put(0, r);
        ws.setMergeData(Collections.singletonList(
                new IRange().setStartRow(0).setStartColumn(0).setEndRow(1).setEndColumn(1)));
        src.getSheets().put("s1", ws);

        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        UniverXlsx.write(src, buf);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(buf.toByteArray()));

        assertThat(back.getSheets().get("s1").getCellData().get(0).get(0).getV()).isEqualTo("Bold");
        assertThat(back.getSheets().get("s1").getMergeData()).hasSize(1);
    }
}
