package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 内边距数据。
 * Padding data (top, bottom, left, right).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IPaddingData extends AbstractUniverModel {
    /** 上内边距 / top padding. */
    private Integer t;
    /** 下内边距 / bottom padding. */
    private Integer b;
    /** 左内边距 / left padding. */
    private Integer l;
    /** 右内边距 / right padding. */
    private Integer r;
}
