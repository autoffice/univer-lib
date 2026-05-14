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

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 共享公式注册器测试 / Shared-formula registry tests.
 */
class SharedFormulaRegistryTest {

    @Test
    void should_group_same_formula_into_one_si() {
        SharedFormulaRegistry r = new SharedFormulaRegistry();
        String s1 = r.registerRead(0, 0, 0, "A$1+1");
        String s2 = r.registerRead(0, 0, 1, "A$1+1");
        assertThat(s1).isEqualTo(s2);
        assertThat(r.masterFormulaOf(s1)).contains("A$1+1");
    }

    @Test
    void should_assign_different_si_to_different_formulas() {
        SharedFormulaRegistry r = new SharedFormulaRegistry();
        String s1 = r.registerRead(0, 0, 0, "A$1+1");
        String s2 = r.registerRead(0, 1, 0, "B$1+1");
        assertThat(s1).isNotEqualTo(s2);
    }

    @Test
    void should_pick_bottom_right_as_master_on_read() {
        SharedFormulaRegistry r = new SharedFormulaRegistry();
        String s1 = r.registerRead(0, 0, 0, "SUM(A1:B1)");
        r.registerRead(0, 2, 3, "SUM(A1:B1)");
        r.registerRead(0, 1, 2, "SUM(A1:B1)");
        int[] master = r.masterCoordOf(s1);
        assertThat(master).containsExactly(2, 3);
    }

    @Test
    void should_not_share_si_across_sheets() {
        SharedFormulaRegistry r = new SharedFormulaRegistry();
        String s1 = r.registerRead(0, 0, 0, "SUM(A1:B1)");
        String s2 = r.registerRead(1, 0, 0, "SUM(A1:B1)");
        assertThat(s1).isNotEqualTo(s2);
        // master 坐标各自独立 / Master coord remains per-si.
        assertThat(r.masterCoordOf(s1)).containsExactly(0, 0);
        assertThat(r.masterCoordOf(s2)).containsExactly(0, 0);
    }

    @Test
    void should_pick_bottom_right_as_master_on_write() {
        SharedFormulaRegistry r = new SharedFormulaRegistry();
        r.registerWrite(0, 0, 0, "si-x", "SUM(A1:B1)");
        r.registerWrite(0, 2, 3, "si-x", "SUM(A1:B1)");
        r.registerWrite(0, 1, 2, "si-x", "SUM(A1:B1)");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("s");
            r.applyOnWorkbook(wb);
            assertThat(wb.getSheetAt(0).getRow(2).getCell(3).getCellFormula()).isEqualTo("SUM(A1:B1)");
            assertThat(wb.getSheetAt(0).getRow(0).getCell(0).getCellFormula()).isEqualTo("SUM(A1:B1)");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void should_strip_leading_equals_sign() {
        SharedFormulaRegistry r = new SharedFormulaRegistry();
        r.registerWrite(0, 0, 0, "si1", "=SUM(A1:B1)");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("s");
            r.applyOnWorkbook(wb);
            assertThat(wb.getSheetAt(0).getRow(0).getCell(0).getCellFormula()).isEqualTo("SUM(A1:B1)");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
