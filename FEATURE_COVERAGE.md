# Univer Sheets ↔ XLSX 转换功能覆盖情况

## ✅ 已实现的功能

### 核心功能 (Core Features)
- ✅ **单元格数据** (Cell Data) - `CellConverter`
  - 文本、数字、布尔值、公式
  - 单元格样式（字体、颜色、边框、对齐等）
  - 数字格式 (numfmt)
- ✅ **富文本** (Rich Text) - `RichTextConverter`
  - 单元格内多样式文本
- ✅ **共享公式** (Shared Formula) - `SharedFormulaRegistry`
- ✅ **样式** (Styles) - `StyleConverter`
  - 字体、填充、边框、对齐、数字格式
- ✅ **行列属性** (Row/Column Data) - `WorksheetConverter`
  - 行高、列宽
  - 隐藏行列
- ✅ **合并单元格** (Merged Cells) - `WorksheetConverter`
- ✅ **冻结窗格** (Freeze Panes) - `WorksheetConverter`
  - 读写 `IFreeze` (xSplit, ySplit, startRow, startColumn)
- ✅ **工作表属性** (Worksheet Attributes) - `WorksheetConverter`
  - 标签颜色 (tabColor)
  - 显示网格线 (showGridlines)
  - 从右到左 (rightToLeft)
  - 缩放比例 (zoomRatio - 仅写)
  - 默认行高列宽
- ✅ **批注/注释** (Comments/Notes) - `CommentConverter`
  - SHEET_NOTE_PLUGIN resource
  - 位置、内容、显示状态、尺寸
- ✅ **条件格式** (Conditional Formatting) - `ConditionalFormattingConverter`
  - SHEET_CONDITIONAL_FORMATTING_PLUGIN resource
  - 规则类型、范围、样式
- ✅ **图片** (Images/Pictures) - `PictureConverter`
  - SHEET_DRAWING_PLUGIN resource
  - base64 data URI、位置、锚点类型

## ❌ 未实现的功能

### 高优先级（Excel 常用功能）

#### 1. **超链接** (Hyperlinks) - `hyper-link.zh-CN.mdx`
- **XLSX 侧**: `XSSFHyperlink` (POI API)
  - 单元格超链接：`cell.getHyperlink()` / `cell.setHyperlink()`
  - 类型：URL、邮件、文档内部、文件
- **Univer 侧**: 富文本内嵌 link
  - `ICellData.p.body.dataStream` 中的 `\r\n` 标记
  - `ICellData.p.body.customRanges` 存储 link 元数据
- **转换难点**: 
  - POI 的 hyperlink 是单元格级别，Univer 是富文本片段级别
  - 需要在 `RichTextConverter` 中处理 link 的读写

#### 2. **数据验证** (Data Validation) - `data-validation.zh-CN.mdx`
- **XLSX 侧**: `XSSFDataValidation` (POI API)
  - `sheet.getDataValidations()` 读取
  - `sheet.addValidationData()` 写入
  - 支持类型：数字、整数、文本长度、日期、列表、自定义公式
- **Univer 侧**: 独立的 plugin resource
  - Plugin: `@univerjs/sheets-data-validation`
  - Resource name: 待确认（可能是 `SHEET_DATA_VALIDATION_PLUGIN`）
  - 数据结构：规则 ID、范围、条件类型、条件值、选项
- **转换难点**:
  - POI 的 DataValidation 结构与 Univer 的规则模型需要映射
  - 下拉列表的来源（固定值 vs 引用范围）需要转换

#### 3. **筛选** (Filter/AutoFilter) - `filter.zh-CN.mdx`
- **XLSX 侧**: `XSSFAutoFilter` (POI API)
  - `sheet.setAutoFilter(range)` 设置筛选范围
  - `CTAutoFilter` 底层 XML 结构存储筛选条件
- **Univer 侧**: 独立的 plugin
  - Plugin: `@univerjs/sheets-filter`
  - 数据结构：筛选范围、列筛选条件
- **转换难点**:
  - POI 对 AutoFilter 的条件读写支持有限
  - 需要解析底层 XML (CTFilterColumn, CTCustomFilters 等)

#### 4. **表格** (Table) - `table.zh-CN.mdx`
- **XLSX 侧**: `XSSFTable` (POI API)
  - `sheet.getTables()` 读取
  - `sheet.createTable()` 创建
  - 表格样式、列定义、筛选按钮
- **Univer 侧**: 独立的 plugin
  - Plugin: `@univerjs/sheets-table`
  - 数据结构：表格范围、样式、列配置
- **转换难点**:
  - Excel Table 是结构化引用，Univer Table 可能是不同的概念
  - 需要确认两者的语义对应关系

#### 5. **定义名称** (Defined Names) - `defined-names.zh-CN.mdx`
- **XLSX 侧**: `XSSFName` (POI API)
  - `workbook.getAllNames()` / `workbook.getName()`
  - `workbook.createName()` 创建
  - 作用域：工作簿级 / 工作表级
- **Univer 侧**: `IWorkbookData.definedNames`
  - 数据结构：name, formulaOrRefString, localSheetId, comment
- **转换难点**:
  - 公式字符串格式可能需要转换（A1 vs R1C1）
  - 作用域映射（全局 vs 局部）

### 中优先级（进阶功能）

#### 6. **图表** (Charts) - `charts.zh-CN.mdx`
- **XLSX 侧**: `XSSFChart` (POI API)
  - `drawing.createChart()` 创建
  - 图表类型、数据源、样式
- **Univer 侧**: 独立的 plugin
  - Plugin: `@univerjs-pro/sheets-chart`（Pro 功能）
  - 数据结构：图表类型、数据范围、配置
- **转换难点**:
  - 图表类型和配置项映射复杂
  - Univer 图表是 Pro 功能，可能需要特殊处理

#### 7. **数据透视表** (Pivot Table) - `pivot-table.zh-CN.mdx`
- **XLSX 侧**: `XSSFPivotTable` (POI API)
  - POI 对透视表的支持有限（主要是读取）
- **Univer 侧**: 独立的 plugin
  - Plugin: `@univerjs-pro/sheets-pivot-table`（Pro 功能）
- **转换难点**:
  - POI 对透视表的写入支持不完善
  - 透视表结构复杂，映射工作量大

#### 8. **迷你图** (Sparkline) - `sparkline.zh-CN.mdx`
- **XLSX 侧**: POI 不直接支持 Sparkline
  - 需要通过底层 XML 操作 (CTSparkline)
- **Univer 侧**: 独立的 plugin
  - Plugin: `@univerjs/sheets-sparkline`
- **转换难点**:
  - POI 无高级 API，需要操作底层 XML
  - 迷你图类型和配置项映射

#### 9. **形状** (Shapes) - `shapes.zh-CN.mdx`
- **XLSX 侧**: `XSSFSimpleShape` (POI API)
  - `drawing.createSimpleShape()` 创建
  - 形状类型、位置、样式
- **Univer 侧**: SHEET_DRAWING_PLUGIN 的一部分
  - `drawingType` 区分图片 vs 形状
  - 当前 `PictureConverter` 只处理了 `drawingType=0`（图片）
- **转换难点**:
  - 需要扩展 `PictureConverter` 支持其他 drawingType
  - 形状的几何属性和样式映射

### 低优先级（特殊功能）

#### 10. **排序** (Sort) - `sort.zh-CN.mdx`
- **说明**: 排序通常是 UI 操作，不存储在文件中
- **XLSX 侧**: 无持久化数据
- **Univer 侧**: 运行时功能，无需转换

#### 11. **查找替换** (Find & Replace) - `find-replace.zh-CN.mdx`
- **说明**: UI 功能，无需转换

#### 12. **打印设置** (Print) - `print.zh-CN.mdx`
- **XLSX 侧**: `XSSFPrintSetup`, `XSSFSheet.getPrintSetup()`
  - 页面设置、页眉页脚、打印区域
- **Univer 侧**: 待确认是否有对应的数据结构
- **转换难点**:
  - 需要确认 Univer 是否支持打印设置的持久化

#### 13. **水印** (Watermark) - `watermark.zh-CN.mdx`
- **说明**: Univer 特有功能，Excel 无直接对应

#### 14. **协同编辑** (Collaboration) - `collaboration.zh-CN.mdx`
- **说明**: 运行时功能，无需转换

#### 15. **编辑历史** (Edit History) - `edit-history.zh-CN.mdx`
- **说明**: 运行时功能，无需转换

## 📊 功能覆盖统计

### 已实现
- ✅ 核心单元格数据（文本、数字、公式、样式）
- ✅ 富文本
- ✅ 共享公式
- ✅ 行列属性（高度、宽度、隐藏）
- ✅ 合并单元格
- ✅ 冻结窗格
- ✅ 工作表属性（标签颜色、网格线等）
- ✅ 批注/注释
- ✅ 条件格式
- ✅ 图片

**已实现功能数**: 10 项

### 未实现（高优先级）
- ❌ 超链接
- ❌ 数据验证
- ❌ 筛选/自动筛选
- ❌ 表格
- ❌ 定义名称

**高优先级缺失**: 5 项

### 未实现（中优先级）
- ❌ 图表
- ❌ 数据透视表
- ❌ 迷你图
- ❌ 形状（非图片的 drawing）

**中优先级缺失**: 4 项

### 未实现（低优先级）
- ❌ 打印设置
- 其他为 UI 功能或 Univer 特有功能

**低优先级缺失**: 1 项

## 🎯 推荐实现顺序

基于 Excel 使用频率和实现难度，推荐按以下顺序补充：

1. **超链接** (Hyperlinks) - 常用功能，实现难度中等
2. **定义名称** (Defined Names) - 常用功能，实现难度低
3. **数据验证** (Data Validation) - 常用功能，实现难度中等
4. **筛选** (Filter) - 常用功能，实现难度中等
5. **表格** (Table) - 进阶功能，实现难度中等
6. **形状** (Shapes) - 扩展现有 PictureConverter
7. **图表** (Charts) - 复杂功能，实现难度高
8. **迷你图** (Sparkline) - 小众功能
9. **数据透视表** (Pivot Table) - 复杂功能，POI 支持有限
10. **打印设置** (Print Setup) - 低频功能

## 📝 实现建议

### 超链接实现思路
1. 在 `RichTextConverter` 中扩展，处理 `customRanges` 中的 link 类型
2. 读路径：`cell.getHyperlink()` → 转为富文本 link
3. 写路径：解析富文本 link → `cell.setHyperlink()`
4. 注意处理单元格级 hyperlink vs 富文本片段级 link 的差异

### 定义名称实现思路
1. 新增 `DefinedNameConverter` 类
2. 读路径：`workbook.getAllNames()` → `IWorkbookData.definedNames`
3. 写路径：`IWorkbookData.definedNames` → `workbook.createName()`
4. 注意作用域映射（全局 vs 工作表级）

### 数据验证实现思路
1. 新增 `DataValidationConverter` 类
2. 读路径：`sheet.getDataValidations()` → plugin resource
3. 写路径：plugin resource → `sheet.addValidationData()`
4. 需要映射验证类型和条件

### 筛选实现思路
1. 新增 `FilterConverter` 类
2. 读路径：`sheet.getCTWorksheet().getAutoFilter()` → plugin resource
3. 写路径：plugin resource → `sheet.setAutoFilter()`
4. 可能需要操作底层 XML 来读写筛选条件

## 🔍 验证方法

对于每个新实现的功能，建议：
1. 创建包含该功能的测试 Excel 文件
2. 编写 round-trip 单元测试
3. 验证前端能正确渲染和交互
4. 检查导出的 Excel 文件在 Excel/WPS 中能正常打开和使用
