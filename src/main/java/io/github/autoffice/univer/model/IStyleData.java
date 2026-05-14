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
 * 单元格样式。
 * Cell style data (maps to Univer IStyleData).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IStyleData extends AbstractUniverModel {
    /** 字体 / font family. */
    private String ff;
    /** 字号(pt) / font size in points. */
    private Integer fs;
    /** 斜体 / italic. */
    private BooleanNumber it;
    /** 粗体 / bold. */
    private BooleanNumber bl;
    /** 下划线 / underline. */
    private ITextDecoration ul;
    /** 删除线 / strikethrough. */
    private ITextDecoration st;
    /** 上划线 / overline. */
    private ITextDecoration ol;
    /** 背景色 / background color. */
    private IColorStyle bg;
    /** 边框 / border. */
    private IBorderData bd;
    /** 字体色 / font color. */
    private IColorStyle cl;
    /** 上下标(1=normal,2=sub,3=super) / vertical align. */
    private Integer va;
    /** 旋转 / text rotation. */
    private ITextRotation tr;
    /** 水平对齐(1=left,2=center,3=right) / horizontal align. */
    private Integer ht;
    /** 垂直对齐(1=top,2=center,3=bottom) / vertical align. */
    private Integer vt;
    /** 溢出策略(1=overflow,2=clip,3=wrap) / wrap strategy. */
    private Integer tb;
    /** 内边距 / padding. */
    private IPaddingData pd;
    /** 数字格式 / number format. */
    private INumfmtLocal n;
}
