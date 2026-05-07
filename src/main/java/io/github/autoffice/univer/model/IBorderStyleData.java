package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 单条边框样式。
 * Single border style entry (line style + color).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IBorderStyleData extends AbstractUniverModel {
    /** 边框线型 / border line style code. */
    private Integer s;
    /** 边框颜色 / border color. */
    private IColorStyle cl;
}
