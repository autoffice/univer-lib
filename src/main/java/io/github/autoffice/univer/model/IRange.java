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
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 区域坐标范围。
 * Range coordinates (start/end row, start/end column, type, sheetId).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IRange extends AbstractUniverModel {
    /** 起始行（0 基） / start row index (0-based). */
    private Integer startRow;
    /** 起始列（0 基） / start column index (0-based). */
    private Integer startColumn;
    /** 结束行（含） / end row index (inclusive). */
    private Integer endRow;
    /** 结束列（含） / end column index (inclusive). */
    private Integer endColumn;
    /** 区域类型 / range type code. */
    private Integer rangeType;
    /** 所属 sheetId / sheetId the range belongs to. */
    private String sheetId;
}
