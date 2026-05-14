# univer-lib Example

演示 `io.github.autoffice:univer-lib` 库的 Web 集成示例：Spring Boot 2.7 后端 + Vue 3 前端，提供 xlsx 导入与导出功能，前端用 Univer Sheets 渲染。

Demo of `io.github.autoffice:univer-lib`: a Spring Boot 2.7 backend + Vue 3 frontend with xlsx import/export, rendered by Univer Sheets.

## Prerequisites / 前置依赖

- JDK 8+（与库本身一致；后端选用 Spring Boot 2.7.18 以保证 JDK 8 兼容）
- Node 18+ / pnpm 或 npm
- 已安装 univer-lib 到本地 Maven 仓库

## Build / 构建步骤

### 1. 安装 univer-lib 到本地仓库

在仓库根目录（`univer-lib/`）执行：

```bash
cd /path/to/univer-lib
./mvnw -q install -DskipTests
```

这会把 `io.github.autoffice:univer-lib:1.0.0` 装进 `~/.m2/repository`，供 example/backend 引用。

### 2. 启动后端

```bash
cd example/backend
../../mvnw spring-boot:run
# 或者：
# mvn spring-boot:run
```

后端默认监听 http://localhost:8080 ，提供两个接口：
- `POST /api/import`（multipart/form-data 字段 `file`）→ 返回 `IWorkbookData` JSON
- `POST /api/export`（请求体为 `IWorkbookData` JSON）→ 返回 xlsx 文件

### 3. 启动前端

```bash
cd example/frontend
npm install   # 或 pnpm install
npm run dev
```

打开 http://localhost:5173 ，左上工具栏可点击 **导入 xlsx** 或 **导出 xlsx**。前端 Vite 已配置 `/api` 代理到 `http://localhost:8080`，无需关心跨域。

## Architecture

```
┌──────────────────────┐  POST /api/import (xlsx)   ┌────────────────────────┐
│  Vue 3 + Univer UI   │ ─────────────────────────▶ │ Spring Boot 2.7 backend│
│   (port 5173)        │ ◀─ IWorkbookData JSON ──── │ univer-lib             │
│                      │                            │  UniverXlsx.read       │
│   univerAPI.create   │                            │                        │
│        Workbook(json)│                            │                        │
│                      │ ─── IWorkbookData JSON ──▶ │  POST /api/export      │
│   fWorkbook.save()   │ ◀───── xlsx bytes ──────── │  UniverXlsx.write      │
└──────────────────────┘                            └────────────────────────┘
```

## Notes

- 前端使用 `@univerjs/presets` + `@univerjs/preset-sheets-core` 渲染。
- 后端 `JsonMapper.get()` 已注册 `IntegerKeyDeserializer`，能正确解析 `cellData` 嵌套的数字字符串键。
