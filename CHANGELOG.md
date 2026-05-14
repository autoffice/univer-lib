# Changelog

本项目所有重要变更会记录在这里。
All notable changes to this project will be documented in this file.

格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本号遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

## [Unreleased]

## [1.0.0] - 2026-05-11

首个公开版本。Initial public release.

### Added

- 公共门面 `io.github.autoffice.univer.UniverXlsx`，静态 `read` / `write` 方法 + `UniverXlsxOptions` 构建器
- xlsx ↔ Univer Sheets `IWorkbookData` JSON 双向转换（JDK 8 兼容）
- **Sidecar 模式**：将完整 `IWorkbookData` JSON 写入自定义 OPC 部件 `/univer/metadata.json`，保证无损往返
- **9 大 sheet 特性**端到端覆盖：单元格数据、样式、合并、富文本、共享公式、行列尺寸、冻结、条件格式、数据透视
- 高级 plugin bridge：`SHEET_DEFINED_NAME_PLUGIN`、`SHEET_FILTER_PLUGIN`、`SHEET_TABLE_PLUGIN`
- `StyleConverter` 样式去重缓存，绕过 POI 64K 单元格样式上限
- `SharedFormulaRegistry` 共享公式 `si` 处理，写入时 master 单元格自动落到右下角
- POJO 通过 `AbstractUniverModel.extras` 实现 Univer 字段前向兼容
- `JsonMapper` + `IntegerKeyDeserializer` 处理 Univer 数字字符串键 (`cellData` / `rowData`)
- 完整 `example/` 演示：Spring Boot 2.7 后端 + Vue 3 + Univer 前端
- 双语（中文 + 英文）公共 API Javadoc
- 测试 70+，覆盖率 ≈47%（POI 兜底分支故意不覆盖）
- Maven profile：`coverage`（JaCoCo）、`lint`（Alibaba p3c-pmd）、`release`（GPG / source / javadoc / Sonatype Central）

### Known limitations

- `UniverXlsxOptions.strictMode` 当前是预留开关，没有 converter 抛 `UniverXlsxUnsupportedFeatureException`
- 长度换算（`LengthUtils`）使用工程近似：1 inch = 96 px = 72 pt，1 char ≈ 7 px
- 外部 xlsx（无 sidecar）读取时，Univer 专属字段（`resources` / `custom` / `appVersion` 等）会缺失，写出时会从默认值兜底

[Unreleased]: https://github.com/autoffice/univer-lib/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/autoffice/univer-lib/releases/tag/v1.0.0
