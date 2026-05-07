package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 行数据（高度 / 自适应高度 / 隐藏）。
 * Row data (height, auto height, hidden flag).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IRowData extends AbstractUniverModel {
    /** 行高 / row height. */
    private Double h;
    /** 自适应行高 / auto-computed row height. */
    private Double ah;
    /** 是否隐藏 / whether the row is hidden. */
    private BooleanNumber hd;
}
