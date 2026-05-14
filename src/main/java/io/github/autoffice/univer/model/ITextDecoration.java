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
 * 文本装饰（下划线、删除线等）。
 * Text decoration (underline / strike-through).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ITextDecoration extends AbstractUniverModel {
    /** 是否显示装饰线 / whether the decoration is shown. */
    private BooleanNumber s;
    /** 是否使用自定义颜色 / whether a custom color is used. */
    private BooleanNumber c;
    /** 装饰线颜色 / decoration line color. */
    private IColorStyle cl;
    /** 装饰线类型 / decoration line type. */
    private Integer t;
}
