package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

/**
 * Univer 中用数字 0/1 表达布尔语义的枚举。
 * Boolean-as-number enum used by Univer model.
 */
public enum BooleanNumber {
    /** 假 / false. */
    FALSE(0),
    /** 真 / true. */
    TRUE(1);

    private final int value;

    BooleanNumber(int value) { this.value = value; }

    /**
     * 获取枚举对应的数字值（0 或 1）。
     * Get the numeric value (0 or 1) of this enum constant.
     */
    @JsonValue
    public int getValue() { return value; }

    /**
     * 根据数字值反序列化为枚举，0 为 FALSE，其余为 TRUE。
     * Deserialize from numeric value: 0 maps to FALSE, any other value maps to TRUE.
     */
    @JsonCreator
    public static BooleanNumber of(int v) { return v == 0 ? FALSE : TRUE; }
}
