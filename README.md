# Excel Certificate Generator

中文 | [English](#english)

## 中文

### 项目简介
`excel-certificate-generator` 是一个基于 Excel VBA 的证件批量生成工具，适用于证书、桌牌、标签等固定版式场景。

- 项目地址：https://github.com/54jayus/excel-certificate-generator
- 开源协议：MIT

### 核心功能
- 选择模板 Sheet 与数据 Sheet
- 选择模板区域（如 `A1:H14`）
- 按数量批量生成模板版式
- 批量写入数据到模板
- 支持按字段分页（可选）

### 环境与兼容性
- 目前已在作者本机的 **Microsoft Office** 与 **WPS** 中可用。
- 尚未系统测试具体 Excel 版本号。

### VBA 宏启用说明
1. 打开 `excel-certificate-generator.xlsm`。
2. 在安全提示中点击“启用内容 / 启用宏”。
3. 如果被安全策略拦截，请将项目放到受信任位置，或在安全中心按需调整宏设置。

> 安全建议：仅对可信来源文件启用宏。

### 快速开始
```bash
git clone https://github.com/54jayus/excel-certificate-generator.git
```

1. 打开 `excel-certificate-generator.xlsm`
2. 在工具界面中完成三步：
   - 第 1 步：选择模板 Sheet 与数据 Sheet
   - 第 2 步：选择模板区域并设置参数
   - 第 3 步：执行生成与写入

### 使用步骤
1. **数据源配置**
   - 模板 Sheet：单份证件样式所在工作表
   - 数据 Sheet：批量数据所在工作表

2. **模板与排版配置**
   - 模板区域：单份模板占用区域（示例：`A1:H14`）
   - 生成数量：生成的模板份数
   - 每行个数：每行排版数量
   - 分页字段：按字段变化自动分页；不需要可选“不分页”

3. **执行操作（顺序重要）**
   - 先点击：`第 1 步：批量生成模板`
   - 再点击：`第 2 步：批量写入数据`

### 项目结构
```text
.
├─ excel-certificate-generator.xlsm
└─ src/
   └─ webui/
      ├─ index.html
      ├─ app.js
      ├─ bridge.js
      └─ styles.css
```

### 许可证
本项目使用 [MIT License](./LICENSE)。

---

## English

### Overview
`excel-certificate-generator` is an Excel VBA tool for batch certificate generation, suitable for fixed-layout scenarios such as certificates, desk signs, and labels.

- Project URL: https://github.com/54jayus/excel-certificate-generator
- License: MIT

### Features
- Select template sheet and data sheet
- Pick template range (e.g. `A1:H14`)
- Batch-generate template layout
- Batch-write data into templates
- Optional pagination by field

### Compatibility
- Confirmed working on the author's local **Microsoft Office** and **WPS** environments.
- Specific Excel versions have not been systematically tested yet.

### Enable VBA Macros
1. Open `excel-certificate-generator.xlsm`.
2. Click **Enable Content / Enable Macros** when prompted.
3. If blocked by security policy, place the file in a trusted location or adjust macro settings in Trust Center as needed.

> Security note: only enable macros for trusted files.

### Quick Start
```bash
git clone https://github.com/54jayus/excel-certificate-generator.git
```

1. Open `excel-certificate-generator.xlsm`
2. Complete the 3 steps in the UI:
   - Step 1: Select template/data sheets
   - Step 2: Configure template range and layout options
   - Step 3: Run generation and data writing

### Usage
1. **Data Source Setup**
   - Template Sheet: worksheet containing the single-certificate layout
   - Data Sheet: worksheet containing batch data

2. **Template & Layout Setup**
   - Template Range: area of one template (example: `A1:H14`)
   - Generate Count: number of templates to generate
   - Per Row Count: templates per row
   - Page Field: paginate when field value changes; choose no pagination if not needed

3. **Execution Order (Important)**
   - First: `Step 1: Batch Generate Template`
   - Then: `Step 2: Batch Write Data`

### Project Structure
```text
.
├─ excel-certificate-generator.xlsm
└─ src/
   └─ webui/
      ├─ index.html
      ├─ app.js
      ├─ bridge.js
      └─ styles.css
```

### License
This project is licensed under the [MIT License](./LICENSE).
