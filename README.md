# excel_datamerger v1.0

[English](#english) | [中文](#中文)

---

<a name="english"></a>

## English Version

### Overview

A powerful Excel/CSV/TXT file merging tool with intelligent column mapping, data quality checking, and smart deduplication features.

### Key Features

- **Multi-format Support**: `.xlsx`, `.xls`, `.csv`, `.txt`
- **Auto-detection**: Multiple worksheets and encoding (UTF-8, GBK, Latin1, etc.)
- **Vertical Stacking**: Concatenate files vertically
- **Source Tracking**: Automatic source file and worksheet columns
- **Column Selection**: Select which columns to exclude from merged output
- **Dark Mode UI**: Optimized for modern interfaces

---

### Quick Start

#### Windows Users

**Method 1: Double-Click (Easiest)**

Simply double-click `run_windows.bat` to automatically install dependencies and launch the program.

**Method 2: Command Line**

```cmd
run_windows.bat
```

**Requirements:**

- Python 3.9 or higher
- Check "Add Python to PATH" during Python installation

**If Python is not installed:**

1. Visit https://www.python.org/downloads/
2. Download and install Python (check "Add Python to PATH")
3. Restart command prompt
4. Double-click `run_windows.bat`

---

#### macOS / Linux Users

**Method 1: Startup Script (Recommended)**

```bash
./run_mac_linux.sh
```

**Method 2: Manual Installation**

```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Launch program
python3 main.py
```

**Note:** macOS users may need to grant execution permission on first run:

```bash
chmod +x run_mac_linux.sh
```

---

### User Guide

#### Basic Workflow

1. **Add Files**

   - Click "Add Files" button to select files to merge
   - Supports multiple selection
   - Supports Excel (.xlsx, .xls), CSV (.csv), TXT (.txt) formats

2. **Select Columns (Optional)**

   - Review all columns in the "Column Selection" panel
   - Check boxes next to columns you want to exclude from the merge
   - Use "Select All", "Deselect All", or "Invert" buttons for quick selection

3. **Configure Options**

   - Check "Normalize Column Names": Use mapping rules to unify column names
   - Check "Enable Fuzzy Matching": Allow partial matches (e.g., "Brand Model" matches "Brand")
   - Check "Remove Duplicates": Delete completely identical rows
   - Check "Smart Deduplication": Deduplicate based on key fields

4. **Start Merging**

   - Click "Start Merge" button
   - Select save location
   - Wait for completion

5. **View Results**
   - Program automatically opens output directory
   - Check log panel for detailed processing information

#### Advanced Features

##### 1. Column Mapping Configuration

Click "⚙️ Column Mapping Config" button to open configuration window:

```json
{
  "Product Barcode": ["Barcode", "UPC", "Item Code"],
  "Product Name": ["Name", "Item Name", "Product"],
  "Brand": ["Brand Name", "Manufacturer"],
  "Quantity": ["Sales Qty", "qty", "Amount"]
}
```

**Configuration Guide:**

- Key (left): Standard column name, used after merging
- Value (right): Alias list, all these names will be mapped to the standard name

**Operations:**

- Click "Save Config" after editing
- Click "Reset to Default" to restore default configuration
- Configuration saved in `column_mappings.json` file

##### 2. Smart Deduplication

**Scenario 1: Remove Completely Identical Rows**

- Check "Remove Duplicates"

**Scenario 2: Key-based Deduplication (Recommended)**

- Check "Smart Deduplication (Key Fields)"
- Enter key fields separated by commas
- Example: `Product Barcode,Date` removes duplicates with same barcode and date

**Example:**

```
Deduplication Key Fields: Product Barcode,Order Date
```

This removes records with identical product barcode AND order date.

##### 3. Column Selection

**Purpose:**

- Remove unwanted columns before merging
- Reduce file size by excluding unnecessary data
- Clean up the output by removing columns you don't need

**How to use:**

1. Add files to the merger
2. In the "Column Selection" panel, you'll see all unique columns from your files
3. Check the boxes next to columns you want to **delete** (note: checked = delete)
4. Use quick actions:
   - "全选" (Select All): Mark all columns for deletion
   - "全不选" (Deselect All): Keep all columns
   - "反选" (Invert): Toggle selection
5. The display shows: `Original Column Name → Mapped Name`
6. Protected columns ("来源文件", "工作表") cannot be deleted

**Note:** Column filtering is applied before merging, so excluded columns won't appear in the output file.

##### 4. Fuzzy Matching

**When to use?**

- Column names vary in format but contain keywords
- Example: "Brand Name", "Brand Code", "Brand Info" all contain "Brand"

**Caution:**

- Fuzzy matching may cause false positives
- Recommend testing without it first, check column mapping report before enabling
- Use longer keywords to reduce false matches

##### 5. View Reports

Two reports are generated during merging:

**Column Mapping Report:**

```
File: sales_data.xlsx-Sheet1
  • Brand Name → Brand [Exact Match]
  • Barcode → Product Barcode [Exact Match]
  • Sales Qty → Quantity [Exact Match]
```

**Data Quality Report:**

```
Total Rows: 1500
Total Columns: 15
Duplicate Rows: 23

Null Values:
  • Notes: 120 rows (8.0%)
  • Customer Contact: 45 rows (3.0%)
```

---

### File Structure

```
excel_datamerger/
├── main.py                    # Program entry point
├── requirements.txt           # Dependencies
├── run_mac_linux.sh          # Startup script (Unix)
├── run_windows.bat           # Startup script (Windows)
├── column_mappings.json      # Column mapping config (auto-generated)
├── excelmerger/              # Core modules
│   ├── __init__.py
│   ├── gui.py               # GUI interface
│   ├── merger.py            # Data processing core
│   ├── io_utils.py          # File I/O
│   ├── logger.py            # Logging
│   └── config_manager.py    # Configuration management
└── logs/                     # Log directory
    └── YYYYMMDD.log         # Date-named logs
```

---

### Use Cases

#### Case 1: Merge Multi-month Sales Data

**Requirements:**

- Multiple monthly sales Excel files
- Slightly different column names (e.g., "Brand" vs "Brand Name")
- Need to remove duplicate orders

**Steps:**

1. Add all sales files
2. Check "Normalize Column Names"
3. Check "Smart Deduplication"
4. Enter deduplication field: `Order Number`
5. Start merge

#### Case 2: Consolidate Inventory Data from Different Sources

**Requirements:**

- Mixed CSV and Excel files
- Inconsistent encoding
- Large column name variations

**Steps:**

1. Open "Column Mapping Config"
2. Add custom mapping rules (e.g., warehouse-related fields)
3. Save configuration
4. Add files and merge
5. Check column mapping report to confirm correct transformation

#### Case 3: Financial Data Aggregation

**Requirements:**

- Financial reports from multiple departments
- Need source tracking
- Check data quality

**Steps:**

1. Add all department files
2. Check "Normalize Column Names"
3. Automatic "Source File" column added after merge
4. Review data quality report
5. Check null value statistics

---

### Technical Details

#### Column Normalization Algorithm

1. **Exact Match** (Highest Priority)

   - Direct match of standard names or aliases (case-insensitive, space/symbol-insensitive)

2. **Fuzzy Match** (Optional, Manual Activation)
   - Match by keyword length (descending order)
   - Minimum length limit (2 characters)
   - Avoid short-word false positives

#### Data Merging Strategy

- Uses `pd.concat()` for vertical stacking
- `join="outer"` preserves all columns
- Auto-fills null values when column names don't match
- Auto-adds "Source File" and "Worksheet" columns for traceability

#### Deduplication Strategy

1. **Full Row Deduplication**

   - Uses `pd.drop_duplicates()`
   - Compares all column values

2. **Smart Deduplication**
   - Based on specified key fields
   - Supports multi-field combinations
   - Keeps first record by default

---

### FAQ

#### Q1: Column names not mapping correctly?

**Solution:**

1. Check "Column Mapping Report" in logs
2. Open "Column Mapping Config"
3. Add missing mapping rules
4. Re-merge

#### Q2: Many empty columns after merge?

**Cause:**

- Large column name variations across files
- Column normalization not enabled

**Solution:**

- Check "Normalize Column Names"
- Configure proper mapping rules
- Enable "Fuzzy Matching" (use with caution)

#### Q3: Too much data lost after deduplication?

**Check:**

1. Review log showing how many rows were deleted
2. Confirm deduplication fields are correct
3. Check duplicate row count in data quality report

**Recommendations:**

- Use smart deduplication instead of full-row deduplication
- Carefully select deduplication key fields

#### Q4: File read failure?

**Possible Causes:**

- Corrupted or incorrectly formatted file
- Encrypted Excel file
- Encoding issues (CSV/TXT)

**Solutions:**

- Check if file can be opened normally
- Try saving as standard format
- Check error messages in logs

#### Q5: Out of memory?

**Cause:**

- Files too large or too many files

**Solutions:**

- Merge in batches
- Close other programs to free memory
- Use a more powerful machine

---

### Tech Stack

- **Python 3.9+**
- **pandas 2.2.0+**: Data processing
- **openpyxl 3.1.2+**: Excel read/write
- **xlrd 2.0.1+**: Legacy Excel support
- **chardet 5.2.0+**: Encoding detection
- **tkinter**: GUI framework

---

### Contributing

Issues and Pull Requests are welcome!

---

### License

MIT License

---

### Contact

For questions or suggestions, please submit an Issue.

If you find this project useful, please consider giving it a star!

---

---

<a name="中文"></a>

## 中文版本

### 概述

一个功能强大的 Excel/CSV/TXT 文件合并工具，支持智能列名映射、数据质量检查、智能去重等高级功能。

![](https://webpic.solo-digitalpass.eu.org/20251106/excel_datamerger.png)

### 主要特性

- **多格式支持**：`.xlsx`, `.xls`, `.csv`, `.txt`
- **自动识别**：多工作表自动识别、自动编码检测（UTF-8, GBK, Latin1 等）
- **纵向堆叠**：纵向堆叠合并
- **来源追溯**：来源文件追溯
- **列选择功能**：可选择要从合并结果中排除的列
- **深色模式**：深色模式界面

---

### 快速开始

#### Windows 用户

**方式 1：双击启动（最简单）**

直接双击 `run_windows.bat` 文件即可自动安装依赖并启动程序。

**方式 2：命令行启动**

```cmd
run_windows.bat
```

**系统要求：**

- Python 3.9 或更高版本
- 安装 Python 时需勾选 "Add Python to PATH"

**如果没有安装 Python：**

1. 访问 https://www.python.org/downloads/
2. 下载并安装 Python（记得勾选"Add Python to PATH"）
3. 重启命令行窗口
4. 双击 `run_windows.bat`

---

#### macOS / Linux 用户

**方式 1：使用启动脚本（推荐）**

```bash
./run_mac_linux.sh
```

**方式 2：手动安装**

```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt

# 启动程序
python3 main.py
```

**注意：** macOS 用户首次运行可能需要给脚本添加执行权限：

```bash
chmod +x run_mac_linux.sh
```

---

### 使用指南

#### 基本操作流程

1. **添加文件**

   - 点击"添加文件"按钮选择要合并的文件
   - 支持多选，可以一次添加多个文件
   - 支持 Excel (.xlsx, .xls), CSV (.csv), TXT (.txt) 格式

2. **选择列（可选）**

   - 在"列选择（勾选要删除的列）"面板中查看所有列
   - 勾选要从合并结果中排除的列
   - 使用"全选"、"全不选"或"反选"按钮快速选择

3. **配置选项**

   - 勾选"统一列名"：使用映射规则统一列名
   - 勾选"启用模糊匹配"：允许部分匹配（例如"品牌型号"匹配"品牌"）
   - 勾选"删除重复行"：删除完全相同的行
   - 勾选"智能去重"：基于关键字段去重

4. **开始合并**

   - 点击"开始合并"按钮
   - 选择保存位置
   - 等待合并完成

5. **查看结果**
   - 程序会自动打开输出文件所在目录
   - 查看日志面板了解详细处理信息

#### 高级功能

##### 1. 列名映射配置

点击"⚙️ 列名映射配置"按钮打开配置窗口：

```json
{
  "商品条码": ["条形码", "条码", "barcode", "UPC"],
  "商品名称": ["名称", "品名", "产品名称"],
  "品牌": ["品牌名称", "brand"],
  "数量": ["销售数量", "qty"]
}
```

**配置说明：**

- 键（左边）：标准列名，合并后统一使用这个名称
- 值（右边）：别名列表，所有这些名称都会被映射为标准列名

**操作：**

- 修改配置后点击"保存配置"
- 点击"重置为默认"恢复默认配置
- 配置保存在 `column_mappings.json` 文件中

##### 2. 智能去重

**场景 1：删除完全相同的行**

- 勾选"删除重复行"

**场景 2：基于关键字段去重（推荐）**

- 勾选"智能去重（基于关键字段）"
- 在输入框中填写关键字段，用逗号分隔
- 例如：`商品条码,日期` 表示根据商品条码和日期组合去重

**示例：**

```
去重关键字段：商品条码,订单日期
```

这将删除商品条码和订单日期都相同的重复记录。

##### 3. 列选择

**功能说明：**

- 在合并前删除不需要的列
- 通过排除不必要的数据来减小文件大小
- 清理输出文件，只保留需要的列

**使用方法：**

1. 添加文件到合并列表
2. 在"列选择（勾选要删除的列）"面板中，会显示所有文件中的唯一列名
3. 勾选要**删除**的列（注意：勾选 = 删除）
4. 使用快捷操作：
   - "全选"：标记所有列为删除
   - "全不选"：保留所有列
   - "反选"：切换选择状态
5. 显示格式为：`原始列名 → 映射后列名`
6. 受保护的列（"来源文件"、"工作表"）无法被删除

**注意：** 列过滤在合并前应用，因此被排除的列不会出现在输出文件中。

##### 4. 模糊匹配

**什么时候使用？**

- 列名格式不统一，但包含关键词
- 例如："品牌名称"、"品牌代码"、"品牌信息" 都包含"品牌"

**注意事项：**

- 模糊匹配可能产生误判
- 建议先不勾选，查看列名映射报告后再决定是否启用
- 使用长关键词可以减少误判

##### 5. 查看报告

合并过程中会生成两个报告：

**列名映射报告：**

```
文件: 销售数据.xlsx-Sheet1
  • 品牌名称 → 品牌 [精确匹配]
  • 条码 → 商品条码 [精确匹配]
  • 销售数量 → 数量 [精确匹配]
```

**数据质量报告：**

```
总行数: 1500
总列数: 15
重复行数: 23

空值情况：
  • 备注: 120 行 (8.0%)
  • 客户联系方式: 45 行 (3.0%)
```

---

### 文件结构

```
excel_datamerger/
├── main.py                    # 程序入口
├── requirements.txt           # 依赖列表
├── run_mac_linux.sh          # 启动脚本（Unix）
├── run_windows.bat           # 启动脚本（Windows）
├── column_mappings.json      # 列名映射配置（首次运行后生成）
├── excelmerger/              # 核心模块
│   ├── __init__.py
│   ├── gui.py               # GUI界面
│   ├── merger.py            # 数据处理核心
│   ├── io_utils.py          # 文件读写
│   ├── logger.py            # 日志管理
│   └── config_manager.py    # 配置管理
└── logs/                     # 日志文件目录
    └── YYYYMMDD.log         # 日期命名的日志
```

---

### 使用场景示例

#### 场景 1：合并多个月的销售数据

**需求：**

- 多个月的销售 Excel 文件
- 列名略有不同（如"品牌"和"品牌名称"）
- 需要删除重复订单

**操作步骤：**

1. 添加所有销售文件
2. 勾选"统一列名"
3. 勾选"智能去重"
4. 填写去重字段：`订单号`
5. 开始合并

#### 场景 2：整合不同来源的库存数据

**需求：**

- CSV 和 Excel 混合
- 编码不统一
- 列名差异大

**操作步骤：**

1. 打开"列名映射配置"
2. 添加自定义映射规则（如仓库相关字段）
3. 保存配置
4. 添加文件并合并
5. 查看列名映射报告确认转换正确

#### 场景 3：财务数据汇总

**需求：**

- 多个部门的财务报表
- 需要标注数据来源
- 检查数据质量

**操作步骤：**

1. 添加所有部门文件
2. 勾选"统一列名"
3. 合并后自动添加"来源文件"列
4. 查看数据质量报告
5. 检查空值情况

---

### 技术说明

#### 列名归一化算法

1. **精确匹配**（优先级最高）

   - 直接匹配标准名或别名（忽略大小写、空格、符号）

2. **模糊匹配**（可选，需手动启用）
   - 按关键词长度倒序匹配
   - 最小长度限制（2 个字符）
   - 避免短词误判

#### 数据合并策略

- 使用 `pd.concat()` 进行纵向堆叠
- `join="outer"` 保留所有列
- 列名不匹配时自动补充空值
- 自动添加"来源文件"和"工作表"列用于追溯

#### 去重策略

1. **全行去重**

   - 使用 `pd.drop_duplicates()`
   - 比较所有列的值

2. **智能去重**
   - 基于指定的关键字段
   - 支持多字段组合
   - 默认保留第一条记录

---

### 常见问题

#### Q1：列名没有正确映射怎么办？

**解决方法：**

1. 查看日志中的"列名映射报告"
2. 打开"列名映射配置"
3. 添加缺失的映射规则
4. 重新合并

#### Q2：合并后有很多空列？

**原因：**

- 不同文件的列名差异较大
- 未启用列名归一化

**解决方法：**

- 勾选"统一列名"
- 配置正确的映射规则
- 启用"模糊匹配"（谨慎使用）

#### Q3：去重后数据少了很多？

**检查：**

1. 查看日志中显示删除了多少行
2. 确认去重字段是否正确
3. 检查数据质量报告中的重复行数

**建议：**

- 使用智能去重而非全行去重
- 仔细选择去重关键字段

#### Q4：文件读取失败？

**可能原因：**

- 文件损坏或格式不正确
- Excel 文件被加密
- 编码问题（CSV/TXT）

**解决方法：**

- 检查文件是否能正常打开
- 尝试另存为标准格式
- 查看日志中的错误信息

#### Q5：内存不足？

**原因：**

- 文件过大或数量过多

**解决方法：**

- 分批合并
- 关闭其他程序释放内存
- 使用更强大的机器

---

### 技术栈

- **Python 3.9+**
- **pandas 2.2.0+**: 数据处理
- **openpyxl 3.1.2+**: Excel 读写
- **xlrd 2.0.1+**: 旧版 Excel 支持
- **chardet 5.2.0+**: 编码检测
- **tkinter**: GUI 界面

---

### 贡献

欢迎提交 Issue 和 Pull Request！

---

### 许可证

MIT License

---

### 联系方式

如有问题或建议，请通过 Issue 反馈。

如果觉得该项目对你有用，请给个星标支持！
