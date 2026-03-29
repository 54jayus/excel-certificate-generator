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
- 目前仅在作者本机的 **WPS**（版本号：12.1.0.25225）中测试运行，其他环境尚未验证。

### VBA 宏启用说明

如果 WPS 无法启用宏，请参考项目目录中的 [`wps启用宏教程`](./wps启用宏教程) 文件夹，或自行搜索「WPS 启用 VBA 宏」解决。

### 快速开始
```bash
git clone https://github.com/54jayus/excel-certificate-generator.git
```

#### 方式一：加载 xlam 插件【推荐】

此方式可在任意 WPS 文件中启动宏，无需每次打开指定文件。

1. 打开任意 Excel 文件，点击顶部菜单栏 **工具** > **加载项**，加载 `excel-certificate-generator.xlam`。
2. 将宏添加到快速访问工具栏：右键点击顶部菜单栏空白处，选择 **自定义命令** > **其他命令**，在分类中选择 **宏**，找到 `excel-certificate-generator.xlam` 并添加。
3. 之后在任意 WPS 文件中，点击快速访问工具栏中的宏图标即可启动宏。

#### 方式二：直接打开 xlsm 文件

1. 打开 `excel-certificate-generator.xlsm`，在安全提示中点击 **启用宏**。
2. 在工具界面中完成三步：
   - 第 1 步：选择模板 Sheet 与数据 Sheet
   - 第 2 步：选择模板区域并设置参数
   - 第 3 步：执行生成与写入

> 注意：此方式仅在该文件打开时有效。

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
├─ excel-certificate-generator.xlam   # 插件文件（推荐加载方式）
├─ excel-certificate-generator.xlsm   # 含宏的工作簿
├─ src/
│  ├─ vba/                            # VBA 源码
│  └─ webui/                          # 工具界面
│     ├─ index.html
│     ├─ app.js
│     ├─ bridge.js
│     ├─ styles.css
│     └─ pic/                         # 界面截图
├─ wps启用宏教程/                      # WPS VBA 启用教程及安装包
│  ├─ 第一种安装方式/
│  └─ 【第二种安装方式】wps.vba.exe
└─ 模板样式/                          # 示例模板
   ├─ 准考证样式.xlsx
   ├─ 台角纸样式.xlsx
   └─ 桌角纸样式.xlsx
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
- Tested only on the author's local **WPS** (version 12.1.0.25225). Other environments have not been verified.

### Enable VBA Macros

If WPS cannot enable macros, refer to the [`wps启用宏教程`](./wps启用宏教程) folder in this project, or search online for "WPS enable VBA macros".

### Quick Start
```bash
git clone https://github.com/54jayus/excel-certificate-generator.git
```

#### Option 1: Load the xlam Add-in (Recommended)

This approach enables the macro in any WPS file without opening a specific workbook each time.

1. Open any Excel file, then go to **Tools** > **Add-ins** in the menu bar and load `excel-certificate-generator.xlam`.
2. Add the macro to the Quick Access Toolbar: right-click on an empty area of the menu bar, select **Customize Quick Access Toolbar** > **More Commands**, choose **Macros** in the category list, and add `excel-certificate-generator.xlam`.
3. The macro icon will appear in the Quick Access Toolbar and can be launched from any WPS file.

#### Option 2: Open the xlsm File Directly

1. Open `excel-certificate-generator.xlsm` and click **Enable Macros** when prompted.
2. Complete the 3 steps in the UI:
   - Step 1: Select template/data sheets
   - Step 2: Configure template range and layout options
   - Step 3: Run generation and data writing

> Note: this approach only works while that specific file is open.

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
├─ excel-certificate-generator.xlam   # Add-in file (recommended)
├─ excel-certificate-generator.xlsm   # Macro-enabled workbook
├─ src/
│  ├─ vba/                            # VBA source code
│  └─ webui/                          # Tool UI
│     ├─ index.html
│     ├─ app.js
│     ├─ bridge.js
│     ├─ styles.css
│     └─ pic/                         # Screenshots
├─ wps启用宏教程/                      # WPS VBA setup guide and installers
│  ├─ 第一种安装方式/
│  └─ 【第二种安装方式】wps.vba.exe
└─ 模板样式/                          # Example templates
   ├─ 准考证样式.xlsx
   ├─ 台角纸样式.xlsx
   └─ 桌角纸样式.xlsx
```

### License
This project is licensed under the [MIT License](./LICENSE).
