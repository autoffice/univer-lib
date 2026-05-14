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

import java.util.List;
import java.util.Map;

/**
 * 富文本文档（Univer Doc 最小可 round-trip 结构）。
 * Rich text document (Univer Doc minimal shape for round-trip).
 */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IDocumentData extends AbstractUniverModel {
    /** 文档 ID / document id. */
    private String id;
    /** 正文 / body. */
    private Body body;
    /** 文档样式 / document style. */
    private Map<String, Object> documentStyle;

    /**
     * 富文本正文。
     * Rich text body.
     */
    @Data
    @NoArgsConstructor
    @Accessors(chain = true)
    @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    public static class Body extends AbstractUniverModel {
        /** 文本流 / data stream. */
        private String dataStream;
        /** 文本 run 列表 / text runs. */
        private List<TextRun> textRuns;
        /** 段落列表 / paragraphs. */
        private List<Paragraph> paragraphs;
    }

    /**
     * 文本 run（带样式区间）。
     * Text run with style range.
     */
    @Data
    @NoArgsConstructor
    @Accessors(chain = true)
    @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    public static class TextRun extends AbstractUniverModel {
        /** 起始位置 / start index. */
        private Integer st;
        /** 结束位置 / end index. */
        private Integer ed;
        /** 文本样式 / text style. */
        private IStyleData ts;
    }

    /**
     * 段落分隔。
     * Paragraph break.
     */
    @Data
    @NoArgsConstructor
    @Accessors(chain = true)
    @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    public static class Paragraph extends AbstractUniverModel {
        /** 起始索引 / start index. */
        private Integer startIndex;
    }
}
