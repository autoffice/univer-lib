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

/**
 * 读取 xlsx 时抛出的异常。
 * Thrown when reading xlsx fails.
 */
public class UniverXlsxReadException extends UniverXlsxException {
    public UniverXlsxReadException(String msg) {
        super(msg);
    }

    public UniverXlsxReadException(String msg, Throwable cause) {
        super(msg, cause);
    }
}
