package io.github.autoffice.univer.util;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.module.SimpleModule;

/**
 * 库内统一的 Jackson ObjectMapper。
 * Shared ObjectMapper for the library.
 */
public final class JsonMapper {
    private static final ObjectMapper MAPPER = new ObjectMapper()
            .setSerializationInclusion(JsonInclude.Include.NON_NULL)
            .disable(SerializationFeature.FAIL_ON_EMPTY_BEANS)
            .disable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES);

    static {
        // 注册 Integer key 反序列化，支持 JSON 对象数字字符串 key 转为 Integer。
        // Register Integer key deserializer for maps keyed by Integer.
        SimpleModule mod = new SimpleModule();
        mod.addKeyDeserializer(Integer.class, new IntegerKeyDeserializer());
        MAPPER.registerModule(mod);
    }

    private JsonMapper() {}

    /**
     * 获取库内共享的 ObjectMapper 实例。
     * Return the shared ObjectMapper instance used across the library.
     */
    public static ObjectMapper get() { return MAPPER; }
}
