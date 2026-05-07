package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 单元格边框数据（上下左右 + 四角对角线）。
 * Cell border data (top, bottom, left, right and four diagonal corners).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IBorderData extends AbstractUniverModel {
    /** 上边框 / top border. */
    private IBorderStyleData t;
    /** 下边框 / bottom border. */
    private IBorderStyleData b;
    /** 左边框 / left border. */
    private IBorderStyleData l;
    /** 右边框 / right border. */
    private IBorderStyleData r;
    /** 左上对角线 / top-left diagonal. */
    private IBorderStyleData tl;
    /** 右上对角线 / top-right diagonal. */
    private IBorderStyleData tr;
    /** 左下对角线 / bottom-left diagonal. */
    private IBorderStyleData bl;
    /** 右下对角线 / bottom-right diagonal. */
    private IBorderStyleData br;
}
