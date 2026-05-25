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

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 库公共异常类型的构造器覆盖。
 * Covers both single-arg and (msg, cause) constructors for each exception type.
 */
class UniverXlsxExceptionsTest {

    @Test
    void should_construct_base_exception_with_message_only() {
        UniverXlsxException e = new UniverXlsxException("boom");
        assertThat(e).isInstanceOf(IOException.class);
        assertThat(e.getMessage()).isEqualTo("boom");
        assertThat(e.getCause()).isNull();
    }

    @Test
    void should_construct_base_exception_with_cause() {
        Throwable cause = new RuntimeException("root");
        UniverXlsxException e = new UniverXlsxException("wrap", cause);
        assertThat(e.getMessage()).isEqualTo("wrap");
        assertThat(e.getCause()).isSameAs(cause);
    }

    @Test
    void should_construct_read_exception_with_both_constructors() {
        UniverXlsxReadException a = new UniverXlsxReadException("read-msg");
        UniverXlsxReadException b = new UniverXlsxReadException("read-msg", new RuntimeException("x"));
        assertThat(a).isInstanceOf(UniverXlsxException.class);
        assertThat(b.getCause()).isInstanceOf(RuntimeException.class);
    }

    @Test
    void should_construct_write_exception_with_both_constructors() {
        UniverXlsxWriteException a = new UniverXlsxWriteException("write-msg");
        UniverXlsxWriteException b = new UniverXlsxWriteException("write-msg", new RuntimeException("y"));
        assertThat(a).isInstanceOf(UniverXlsxException.class);
        assertThat(b.getCause()).isInstanceOf(RuntimeException.class);
    }

    @Test
    void should_construct_unsupported_feature_exception() {
        UniverXlsxUnsupportedFeatureException e = new UniverXlsxUnsupportedFeatureException("feat");
        assertThat(e).isInstanceOf(UniverXlsxException.class);
        assertThat(e.getMessage()).isEqualTo("feat");
    }
}
