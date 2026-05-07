package io.github.autoffice.univer.util;

import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.KeyDeserializer;

/**
 * 把 JSON 对象中的数字字符串 key 反序列化为 Integer。
 * Deserialize numeric string keys to Integer.
 */
public class IntegerKeyDeserializer extends KeyDeserializer {
    @Override
    public Object deserializeKey(String key, DeserializationContext ctxt) {
        return Integer.parseInt(key);
    }
}
