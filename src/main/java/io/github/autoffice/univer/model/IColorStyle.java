package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 颜色样式。
 * Color style (rgb or theme color).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IColorStyle extends AbstractUniverModel {
    /** RGB 颜色（#rrggbb） / rgb color hex. */
    private String rgb;
    /** 主题色索引 / theme color index. */
    private Integer th;
}
