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

import java.util.Map;

/**
 * 单元格数据。
 * Cell data (maps to Univer ICellData).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ICellData extends AbstractUniverModel {
    /** 原始值（字符串/数字/布尔） / raw value. */
    private Object v;
    /** 样式 id 或样式对象 / style id (String) or inline IStyleData. */
    private Object s;
    /** 类型 / cell value type. */
    private CellValueType t;
    /** 富文本 / rich text. */
    private IDocumentData p;
    /** 公式 / formula. */
    private String f;
    /** 公式 id / shared formula id. */
    private String si;
    /** 自定义字段 / custom payload. */
    private Map<String, Object> custom;
}
