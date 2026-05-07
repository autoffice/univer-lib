package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonAnyGetter;
import com.fasterxml.jackson.annotation.JsonAnySetter;
import com.fasterxml.jackson.annotation.JsonIgnore;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

/**
 * 所有 Univer POJO 的基类，承接未知字段以防版本升级丢数据。
 * Base class that captures unknown fields via extras map for forward compatibility.
 */
public abstract class AbstractUniverModel {
    @JsonIgnore
    private final Map<String, Object> extras = new LinkedHashMap<>();

    /**
     * 获取承接未知字段的 extras 映射，供 Jackson 序列化时平铺输出。
     * Return the extras map that captures unknown fields; Jackson flattens it during serialization.
     */
    @JsonAnyGetter
    public Map<String, Object> getExtras() { return extras; }

    /**
     * 反序列化时回收未识别字段到 extras，避免版本升级丢失数据。
     * Collect unrecognized fields into extras during deserialization to preserve forward compatibility.
     */
    @JsonAnySetter
    public void putExtra(String key, Object value) { extras.put(key, value); }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        return Objects.equals(extras, ((AbstractUniverModel) o).extras);
    }

    @Override
    public int hashCode() { return Objects.hash(extras); }
}
