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

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 工作簿数据（根快照）。
 * Workbook data (root snapshot, maps to Univer IWorkbookData).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IWorkbookData extends AbstractUniverModel {
    /** 工作簿 ID / workbook id. */
    private String id;
    /** 工作簿名称 / workbook name. */
    private String name;
    /** 应用版本 / app version. */
    private String appVersion;
    /** 区域设置 / locale. */
    private String locale;
    /** 样式引用表 / styles map. */
    private Map<String, IStyleData> styles = new LinkedHashMap<>();
    /** 工作表顺序 / sheet order. */
    private List<String> sheetOrder;
    /** 工作表集合 / sheets map. */
    private Map<String, IWorksheetData> sheets = new LinkedHashMap<>();
    /** 插件资源（不透明 JSON） / plugin resources (opaque JSON). */
    private Object resources;
}
