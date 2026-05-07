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

    @JsonValue
    public int getValue() { return value; }

    @JsonCreator
    public static BooleanNumber of(int v) { return v == 0 ? FALSE : TRUE; }
}
