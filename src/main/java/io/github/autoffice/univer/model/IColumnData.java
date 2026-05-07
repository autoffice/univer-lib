package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 列数据（宽度 / 隐藏）。
 * Column data (width, hidden flag).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IColumnData extends AbstractUniverModel {
    /** 列宽 / column width. */
    private Double w;
    /** 是否隐藏 / whether the column is hidden. */
    private BooleanNumber hd;
}
