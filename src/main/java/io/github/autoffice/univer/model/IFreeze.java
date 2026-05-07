package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 冻结窗格配置。
 * Freeze pane configuration (split point + start row/column).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IFreeze extends AbstractUniverModel {
    /** 起始行 / start row index. */
    private Integer startRow;
    /** 起始列 / start column index. */
    private Integer startColumn;
    /** 水平冻结列数 / number of frozen columns (x-split). */
    private Integer xSplit;
    /** 垂直冻结行数 / number of frozen rows (y-split). */
    private Integer ySplit;
}
