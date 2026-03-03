# 智能表格数据填充工具

一个功能强大的表格数据自动抓取和填充应用，支持多种文件格式和智能匹配。

## 功能特性

### 📁 多格式支持
- **Excel**: `.xlsx`, `.xls`
- **Word**: `.docx`, `.doc`
- **WPS**: `.docx`, `.xlsx` (通过标准格式支持)
- **CSV**: `.csv`
- **文本**: `.txt`

### 🧠 智能识别
- **自动表格识别**: 自动检测文档中的所有表格，无论表格位置
- **同义词匹配**: 智能识别相似字段名称
  - "总资产" ↔ "资产总额" ↔ "资产合计"
  - "净利润" ↔ "净利" ↔ "纯利润"
  - "营业收入" ↔ "营收" ↔ "销售收入"
  - 等等（可扩展）
- **多格式日期识别**:
  - `2025/9`, `2025-9`, `2025.9`
  - `2025年9月`
  - `202509`

### 🔄 智能换算
- **单位识别**: 自动识别文档中的单位标注（"单位：亿元/万元 %"）
- **单位换算**: 自动进行单位转换
  - 亿元 ↔ 万元 ↔ 元
  - 百分比处理（根据表格外标注自动加/减 %）
- **精确计算**: 支持小数点后4位精度

### 📊 数据处理
- **智能填充**: 只填充空白单元格，保留已有数据
- **批量处理**: 支持多表格同时处理
- **匹配统计**: 提供填充率和转换统计信息

## 使用方法

### 1. 上传文件
- **文件A（数据源文件）**: 上传包含完整数据的文件
- **文件B（数据缺失文件）**: 上传需要填充数据的文件

### 2. 文件要求
- 文件B中表格的横轴应为匹配字段
- 纵轴应为时间点
- 表格可以在文档的任意位置
- 确保文件A中包含文件B缺失数据的匹配项

### 3. 处理流程
1. 点击"开始处理"按钮
2. 系统自动识别表格结构和单位
3. 智能匹配字段和日期
4. 自动填充空白单元格
5. 进行单位换算（如需要）

### 4. 下载结果
- 处理完成后，点击"下载结果"按钮
- 结果文件以 Excel 格式保存
- 文件名格式：`填充结果_原文件名.xlsx`

## 技术架构

### 前端
- **框架**: Next.js 16 (App Router)
- **UI组件**: shadcn/ui (基于 Radix UI)
- **样式**: Tailwind CSS 4
- **语言**: TypeScript 5

### 后端
- **文件解析**:
  - Excel: `xlsx` 库
  - Word: `mammoth` 库
  - CSV: `papaparse` 库
- **数据处理**: 自定义智能匹配算法
- **API**: Next.js API Routes

### 核心模块

#### 1. FileParser (`src/lib/file-parser.ts`)
负责解析各种格式的文件并提取表格数据：
```typescript
const parseResult = await FileParser.parseFile(file);
// 返回: { tables, unit, filename }
```

#### 2. DataMatcher (`src/lib/data-matcher.ts`)
智能匹配器，负责数据匹配和填充：
```typescript
const matcher = new DataMatcher(sourceTable, targetTable, ...);
const { filledTable, matchResults } = matcher.matchAndFill();
```

#### 3. BatchDataMatcher (`src/lib/data-matcher.ts`)
批量匹配器，支持多表格同时处理：
```typescript
const matcher = new BatchDataMatcher(sourceTables, targetTables, ...);
const { results } = matcher.matchAll();
```

#### 4. 数据工具 (`src/lib/data-utils.ts`)
提供各种数据处理工具：
```typescript
- normalizeDate() // 日期标准化
- fieldMatches() // 字段匹配（同义词）
- convertUnit() // 单位换算
- extractUnitInfo() // 单位识别
```

## API 接口

### POST /api/process
处理上传的文件并填充数据

**请求**:
- Content-Type: `multipart/form-data`
- Body:
  - `fileA`: 数据源文件
  - `fileB`: 数据缺失文件

**响应**:
```json
{
  "success": true,
  "fileId": "1234567890_filename.xlsx",
  "message": "处理完成",
  "statistics": {
    "totalFilled": 25,
    "totalConverted": 10,
    "tableCount": 2
  }
}
```

### GET /api/download?fileId={fileId}
下载处理后的文件

**参数**:
- `fileId`: 处理后文件的ID

**响应**:
- Content-Type: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- Body: Excel 文件二进制数据

## 扩展功能

### 添加自定义同义词
```typescript
import { addSynonym } from '@/lib/data-utils';

addSynonym('总资产', '总资');
addSynonym('营业收入', '销售额');
```

### 添加自定义单位
```typescript
import { addUnitRule } from '@/lib/data-utils';

addUnitRule('千元', 1000);
addUnitRule('百万元', 1000000);
```

## 注意事项

1. **文件格式**: 确保上传的文件格式正确，损坏的文件可能导致解析失败
2. **数据匹配**: 字段名称越相似，匹配准确率越高
3. **单位标注**: 确保单位标注格式正确（"单位：亿元" 或 "单位：万元 %"）
4. **日期格式**: 支持常见日期格式，特殊格式可能需要扩展
5. **文件大小**: 大文件处理可能需要较长时间

## 错误处理

系统会捕获并显示以下错误：
- 文件格式不支持
- 未找到表格数据
- 处理失败
- 下载失败

## 开发环境

### 安装依赖
```bash
pnpm install
```

### 运行开发服务器
```bash
coze dev
```
服务将在 http://localhost:5000 启动

### 构建生产版本
```bash
coze build
```

### 启动生产服务器
```bash
coze start
```

## 项目结构

```
src/
├── app/
│   ├── api/
│   │   ├── process/
│   │   │   └── route.ts          # 文件处理 API
│   │   └── download/
│   │       └── route.ts          # 文件下载 API
│   ├── page.tsx                  # 主页面
│   ├── layout.tsx                # 布局
│   └── globals.css               # 全局样式
├── components/
│   └── ui/                       # shadcn/ui 组件
├── lib/
│   ├── types.ts                  # 类型定义
│   ├── data-utils.ts             # 数据工具函数
│   ├── data-matcher.ts           # 数据匹配器
│   ├── file-parser.ts            # 文件解析器
│   └── utils.ts                  # 通用工具
└── hooks/
    └── use-mobile.ts             # 移动端检测
```

## 许可证

MIT License
