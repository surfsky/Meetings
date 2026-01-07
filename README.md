# 会议历汇总工具 (Meeting History Summary Tool)

这是一个基于 Vue 3 + TypeScript + Vite 开发的会议记录汇总工具。它可以解析上传的 Excel 文件（包含会议记录、日期、图片等），按周和部门自动分组整理，并支持生成 PDF 和 Word 格式的汇总报表。

## ✨ 主要功能

-   **Excel 解析**: 支持 `.xlsx` 和 `.xls` 格式，自动识别会议日期、部门、内容及嵌入的图片（支持 `DISPIMG` 等特殊图片格式）。
-   **智能分组**: 按年份、月份和周数（1-53周）自动整理会议记录，缺失的周会留空，保证时间线的完整性。
-   **自定义排序**: 支持通过 Excel 中的 "排序" sheet 自定义部门显示的顺序。
-   **预览与交互**:
    -   Web 端表格预览，支持点击图片放大查看。
    -   清晰的表格布局，包含页眉（月份、周、部门）。
-   **多格式导出**:
    -   **PDF 导出**: 自动分页，支持横向布局，包含页码和页边距。
    -   **Word 导出**: 生成格式完美的 `.docx` 文档，图片与文字紧凑排版，支持打印。

## 🛠️ 技术栈

-   **Frontend Framework**: Vue 3 (Script Setup)
-   **Language**: TypeScript
-   **Build Tool**: Vite
-   **UI Library**: Element Plus
-   **Excel Processing**: ExcelJS, JSZip (用于解析特殊图片)
-   **PDF Generation**: jspdf, html2canvas
-   **Word Generation**: docx, file-saver
-   **Date Handling**: Day.js

## 🚀 快速开始

### 环境要求

-   Node.js (推荐 v18+ 或 v20+)
-   npm (v9+)

### 安装依赖

```bash
npm install
```

### 开发模式

启动本地开发服务器：

```bash
npm run dev
```

### 构建生产版本

构建用于生产环境的静态文件（输出到 `dist` 目录）：

```bash
npm run build
```

### 预览生产构建

```bash
npm run preview
```

## 📖 使用说明

1.  **下载模版**: 点击页面上的 "示例下载" 链接，获取标准的 Excel 记录模版。
2.  **填写记录**: 在 Excel 中按列填写会议日期、部门、内容，并插入现场照片。
    -   可在 "排序" sheet 中定义部门的显示顺序。
3.  **上传文件**: 将填写好的 Excel 文件拖拽到上传区域。
4.  **查看预览**: 解析成功后，页面将展示会议记录表格。
5.  **导出报表**:
    -   点击 "下载/打印 PDF" 生成 PDF 文件。
    -   点击 "下载 Word" 生成 Word 文档。

## 📄 License

[MIT](LICENSE)
