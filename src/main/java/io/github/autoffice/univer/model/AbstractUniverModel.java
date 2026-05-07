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

    @JsonAnyGetter
    public Map<String, Object> getExtras() { return extras; }

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
