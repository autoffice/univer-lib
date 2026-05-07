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
