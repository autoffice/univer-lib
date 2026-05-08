package io.github.autoffice.example.controller;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.UniverXlsxException;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.util.JsonMapper;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Collections;

/**
 * xlsx 导入导出 REST 接口。
 * REST endpoints for xlsx import/export.
 */
@RestController
@RequestMapping("/api")
public class UniverXlsxController {

    private final ObjectMapper mapper = JsonMapper.get();

    /**
     * 上传 xlsx 文件并返回 IWorkbookData JSON。
     * Upload xlsx and return IWorkbookData JSON.
     */
    @PostMapping(value = "/import", consumes = MediaType.MULTIPART_FORM_DATA_VALUE,
        produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<String> importXlsx(@RequestParam("file") MultipartFile file) throws Exception {
        if (file.isEmpty()) {
            return errorJson("empty file");
        }
        String filename = file.getOriginalFilename();
        if (filename == null || !filename.toLowerCase().endsWith(".xlsx")) {
            return errorJson("only .xlsx files are supported");
        }
        try (InputStream in = file.getInputStream()) {
            IWorkbookData wb = UniverXlsx.read(in);
            if (wb.getId() == null) {
                wb.setId("imported-" + System.currentTimeMillis());
            }
            String json = mapper.writeValueAsString(wb);
            return ResponseEntity.ok().contentType(MediaType.APPLICATION_JSON).body(json);
        } catch (UniverXlsxException | JsonProcessingException e) {
            return errorJson(e.getMessage());
        }
    }

    /**
     * 接收 IWorkbookData JSON 返回 xlsx 二进制。
     * Accept IWorkbookData JSON and return xlsx bytes.
     */
    @PostMapping(value = "/export", consumes = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<?> exportXlsx(@RequestBody String body,
                                        @RequestParam(value = "name", defaultValue = "export") String name)
            throws Exception {
        try {
            IWorkbookData wb = mapper.readValue(body, IWorkbookData.class);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            UniverXlsx.write(wb, out);
            byte[] bytes = out.toByteArray();
            String fileName = URLEncoder.encode(name + ".xlsx", StandardCharsets.UTF_8.name()).replace("+", "%20");
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.set(HttpHeaders.CONTENT_DISPOSITION,
                "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + fileName);
            return ResponseEntity.ok().headers(headers).body(bytes);
        } catch (JsonProcessingException e) {
            return ResponseEntity.badRequest()
                .contentType(MediaType.APPLICATION_JSON)
                .body(mapper.writeValueAsString(Collections.singletonMap("error", e.getMessage())));
        }
    }

    /** 构造标准 JSON 错误响应 / Build a standard JSON error response. */
    private ResponseEntity<String> errorJson(String message) throws JsonProcessingException {
        return ResponseEntity.badRequest()
            .contentType(MediaType.APPLICATION_JSON)
            .body(mapper.writeValueAsString(
                Collections.singletonMap("error", message == null ? "unknown error" : message)));
    }
}
