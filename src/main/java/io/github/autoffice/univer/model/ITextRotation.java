package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 文本旋转。
 * Text rotation settings (angle + vertical mode).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ITextRotation extends AbstractUniverModel {
    /** 旋转角度（度） / rotation angle in degrees. */
    private Integer a;
    /** 是否竖排 / whether text is vertical. */
    private BooleanNumber v;
}
