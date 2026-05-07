package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

/**
 * 单元格值类型。
 * Cell value type (Univer CellValueType).
 */
public enum CellValueType {
    STRING(1), NUMBER(2), BOOLEAN(3), FORCE_TEXT(4);

    private final int value;
    CellValueType(int v) { this.value = v; }
    @JsonValue public int getValue() { return value; }
    @JsonCreator public static CellValueType of(int v) {
        for (CellValueType t : values()) { if (t.value == v) return t; }
        throw new IllegalArgumentException("Unknown CellValueType: " + v);
    }
}
