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
package io.github.autoffice.univer;

import lombok.Builder;
import lombok.Getter;

/**
 * 读写选项。
 * Reader/writer options.
 */
@Getter
@Builder(toBuilder = true)
public class UniverXlsxOptions {
    /** 严格模式：遇到不支持特性时抛异常。/ Strict mode: throw on unsupported features. */
    @Builder.Default
    private boolean strictMode = false;

    /** 是否写入边车元数据。/ Whether to write the sidecar metadata part. */
    @Builder.Default
    private boolean writeSidecar = true;

    /** 是否美化 JSON。/ Pretty-print sidecar json. */
    @Builder.Default
    private boolean prettyJson = false;

    /** 缺省 locale（无边车时使用）。/ Fallback locale when sidecar missing. */
    @Builder.Default
    private String locale = "enUS";

    /**
     * 获取默认配置。
     * Get default options instance.
     */
    public static UniverXlsxOptions defaults() {
        return UniverXlsxOptions.builder().build();
    }
}
