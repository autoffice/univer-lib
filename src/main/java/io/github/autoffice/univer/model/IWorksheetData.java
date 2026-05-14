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
import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import io.github.autoffice.univer.util.IntegerKeyDeserializer;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 工作表数据。
 * Worksheet data (maps to Univer IWorksheetData).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IWorksheetData extends AbstractUniverModel {
    /** 工作表 ID / worksheet id. */
    private String id;
    /** 工作表名称 / worksheet name. */
    private String name;
    /** 标签颜色 / tab color. */
    private String tabColor;
    /** 是否隐藏 / hidden. */
    private BooleanNumber hidden;
    /** 冻结设置 / freeze pane settings. */
    private IFreeze freeze;
    /** 总行数 / total row count. */
    private Integer rowCount;
    /** 总列数 / total column count. */
    private Integer columnCount;
    /** 默认列宽(px) / default column width in px. */
    private Double defaultColumnWidth;
    /** 默认行高(px) / default row height in px. */
    private Double defaultRowHeight;
    /** 合并区域 / merged cell ranges. */
    private List<IRange> mergeData;

    /** 单元格数据矩阵 / cell data matrix. */
    @JsonDeserialize(keyUsing = IntegerKeyDeserializer.class)
    private Map<Integer, Map<Integer, ICellData>> cellData = new LinkedHashMap<>();

    /** 行数据 / row data. */
    @JsonDeserialize(keyUsing = IntegerKeyDeserializer.class)
    private Map<Integer, IRowData> rowData = new LinkedHashMap<>();

    /** 列数据 / column data. */
    @JsonDeserialize(keyUsing = IntegerKeyDeserializer.class)
    private Map<Integer, IColumnData> columnData = new LinkedHashMap<>();

    /** 行头配置 / row header config. */
    private Header rowHeader;
    /** 列头配置 / column header config. */
    private Header columnHeader;
    /** 是否显示网格线 / show gridlines. */
    private BooleanNumber showGridlines;
    /** 是否从右到左 / right to left. */
    private BooleanNumber rightToLeft;
    /** 默认样式 / default style. */
    private IStyleData defaultStyle;
    /** 缩放比例 / zoom ratio. */
    private Double zoomRatio;
    /** 滚动位置(上) / scroll top. */
    private Double scrollTop;
    /** 滚动位置(左) / scroll left. */
    private Double scrollLeft;

    /**
     * 行/列头配置。
     * Row/column header configuration.
     */
    @Data
    @NoArgsConstructor
    @Accessors(chain = true)
    @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    public static class Header extends AbstractUniverModel {
        /** 宽度 / width. */
        private Double width;
        /** 高度 / height. */
        private Double height;
        /** 是否隐藏 / hidden. */
        private BooleanNumber hidden;
    }
}
