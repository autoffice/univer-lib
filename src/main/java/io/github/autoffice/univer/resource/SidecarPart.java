package io.github.autoffice.univer.resource;

import com.fasterxml.jackson.databind.ObjectWriter;
import io.github.autoffice.univer.UniverXlsxWriteException;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Optional;

/**
 * OPC 边车分区读写器。
 * OPC sidecar part reader/writer for /univer/metadata.json.
 * <p>
 * 存储完整的 IWorkbookData JSON，用于 Univer 特有字段的无损往返。
 */
public final class SidecarPart {
    /** 边车分区路径。 */
    public static final String PART_NAME = "/univer/metadata.json";
    /** 分区内容类型。 */
    public static final String CONTENT_TYPE = "application/json";
    /** 分区关系类型。 */
    public static final String REL_TYPE = "http://schemas.autoffice.io/univer/2026/metadata";

    private SidecarPart() {}

    /** 从 OPC 包中读取边车数据。/ Read sidecar from OPC package. */
    public static Optional<IWorkbookData> read(OPCPackage pkg) {
        try {
            PackagePartName partName = PackagingURIHelper.createPartName(PART_NAME);
            if (!pkg.containPart(partName)) {
                return Optional.empty();
            }
            PackagePart part = pkg.getPart(partName);
            try (InputStream is = part.getInputStream()) {
                IWorkbookData data = JsonMapper.get().readValue(is, IWorkbookData.class);
                return Optional.ofNullable(data);
            }
        } catch (Exception e) {
            // 边车损坏时视为不存在
            return Optional.empty();
        }
    }

    /** 将 IWorkbookData 写入 OPC 包的边车分区。/ Write IWorkbookData to sidecar part in OPC package. */
    public static void write(OPCPackage pkg, IWorkbookData wb, boolean pretty) throws UniverXlsxWriteException {
        try {
            PackagePartName partName = PackagingURIHelper.createPartName(PART_NAME);
            // 覆盖时先删除旧分区再重建，避免 POI 输出流未截断旧内容
            if (pkg.containPart(partName)) {
                pkg.removePart(partName);
            }
            PackagePart part = pkg.createPart(partName, CONTENT_TYPE);
            pkg.addRelationship(partName, TargetMode.INTERNAL, REL_TYPE);
            ObjectWriter writer = pretty
                    ? JsonMapper.get().writerWithDefaultPrettyPrinter()
                    : JsonMapper.get().writer();
            try (OutputStream os = part.getOutputStream()) {
                writer.writeValue(os, wb);
            }
        } catch (Exception e) {
            throw new UniverXlsxWriteException("Failed to write sidecar part: " + e.getMessage(), e);
        }
    }
}
