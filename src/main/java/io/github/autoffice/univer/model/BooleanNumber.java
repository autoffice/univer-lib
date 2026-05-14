/*
 * Copyright © 2026 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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
