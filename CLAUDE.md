# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**excel_webdatamerger** is a desktop GUI application for merging multiple Excel/CSV/TXT files with intelligent column name mapping, data quality checking, and smart deduplication features.

**Tech Stack:**
- Python 3.9+ with tkinter for GUI
- pandas for data processing
- openpyxl/xlrd for Excel file handling
- chardet for encoding detection

## Development Commands

### Running the Application

**macOS/Linux:**
```bash
./run_mac_linux.sh
```

**Windows:**
```cmd
run_windows.bat
```

**Manual (with venv):**
```bash
# Create and activate virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run application
python3 main.py
```

### Dependency Management

All dependencies are managed via `requirements.txt`:
```bash
pip install -r requirements.txt
```

Core dependencies:
- pandas>=2.2.0
- openpyxl>=3.1.2
- xlrd>=2.0.1
- chardet>=5.2.0

## Architecture

### Module Structure

```
excelmerger/
├── gui.py              # Main GUI application (ExcelMergerGUI)
├── merger.py           # Core data processing logic (ExcelMergerCore)
├── io_utils.py         # File I/O utilities (read_file, save_to_excel)
├── config_manager.py   # Column mapping configuration (ConfigManager)
└── logger.py           # Logging setup (setup_logger)
```

### Key Design Patterns

**1. Configuration-Driven Column Mapping**
- User-defined column name mappings stored in `column_mappings.json`
- ConfigManager loads/saves mapping rules
- ExcelMergerCore uses these rules for column normalization

**2. Multi-Stage Column Normalization**
- Stage 1: Exact match (case-insensitive, whitespace-normalized)
- Stage 2: Fuzzy match (optional, keyword-based)
- Stage 3: Preserve original if no match

**3. Data Processing Pipeline**
```
File Selection → Read Files → Normalize Columns → Merge (concat)
→ Deduplicate (optional) → Add Source Tracking → Save → Report
```

**4. Separation of Concerns**
- GUI (gui.py): User interaction, file selection, progress display
- Core Logic (merger.py): Column normalization, data validation, deduplication
- I/O (io_utils.py): Multi-format file reading with encoding auto-detection
- Config (config_manager.py): Mapping rule persistence

### Critical Implementation Details

**Column Normalization (merger.py)**
- `normalize_text()`: Removes whitespace, lowercases, standardizes text
- `ExcelMergerCore.normalize_columns()`: Maps columns using exact/fuzzy matching
- `_build_alias_map()`: Creates lookup dict from config for O(1) matching
- Duplicate column names get suffixed with `_1`, `_2`, etc.

**File Reading Strategy (io_utils.py)**
- Excel files: Try openpyxl first, fallback to xlrd for legacy formats
- CSV/TXT: Auto-detect encoding by trying utf-8-sig, utf-8, gbk, latin1
- Multi-sheet Excel: Returns dict of {sheet_name: DataFrame}
- Always use `sheet_name=None` to read all sheets

**Data Merging**
- Uses `pd.concat()` with `join="outer"` to preserve all columns
- Adds "来源文件" (Source File) and "工作表" (Worksheet) columns for traceability
- Empty cells filled automatically when columns don't align

**Deduplication Strategies**
- Full row dedup: `df.drop_duplicates()` on all columns
- Smart dedup: `df.drop_duplicates(subset=key_columns)` on user-specified key fields

### State Management

**GUI State (gui.py)**
- `self.file_paths`: List of file paths selected by user
- `self.config_manager`: ConfigManager instance
- Boolean options: `normalize_columns`, `remove_duplicates`, `enable_fuzzy_match`, `smart_dedup`
- String options: `dedup_keys` (comma-separated key field names)

**Configuration State (config_manager.py)**
- Loads from `column_mappings.json` on startup
- Falls back to `DEFAULT_MAPPINGS` if file doesn't exist
- Auto-saves when user modifies mappings via GUI

### Threading Model

- File merging runs in a separate thread (started in `gui.py`)
- GUI progress updates via `self.progress_var` and `self.status_text`
- Prevents UI freezing during large file processing

## Common Development Tasks

### Adding a New Column Mapping Rule

1. Edit `column_mappings.json` or use GUI "⚙️ 列名映射配置" button
2. Format: `"标准列名": ["别名1", "别名2", ...]`
3. Changes take effect immediately (no restart needed if using GUI)

### Modifying Column Normalization Logic

Edit `excelmerger/merger.py`:
- `normalize_text()`: Change text normalization rules
- `ExcelMergerCore._fuzzy_match()`: Adjust fuzzy matching algorithm
- Update `_ensure_unique_columns()` for duplicate column handling

### Changing Supported File Formats

Edit `excelmerger/io_utils.py`:
- Add file extension to `read_file()` function
- Implement reader with pandas (e.g., `pd.read_parquet()`)
- Handle encoding/format-specific edge cases

### Adjusting GUI Layout

Edit `excelmerger/gui.py`:
- `_build_ui()`: Main layout construction
- Color scheme defined in `__init__()` via `root.option_add()`
- Dark mode optimized for macOS with specific color values

## Logging

Logs are written to `logs/YYYYMMDD.log` with date-based rotation.

Log locations:
- Daily log files in `logs/` directory
- Console output (stdout) for immediate feedback

The logger captures:
- File read/write operations
- Column mapping transformations
- Data quality statistics
- Error messages with context

## Data Quality Reporting

The application generates two types of reports:

**Column Mapping Report:**
- Shows original → normalized column names
- Indicates match type (精确匹配/模糊匹配/未映射)
- Helps debug mapping issues

**Data Quality Report:**
- Total rows/columns
- Null value statistics per column
- Duplicate row count
- Data type summary

Both reports are displayed in GUI log panel and written to log files.

## Configuration Files

**column_mappings.json** (root directory)
- User-editable column mapping rules
- JSON format: `{"标准列名": ["别名1", "别名2", ...]}`
- Auto-generated with defaults on first run

## Important Constraints

**Memory Limitations:**
- All data loaded into memory via pandas
- Large files (>500MB) may cause memory issues
- Consider chunking for very large datasets

**File Format Support:**
- Excel: .xlsx (openpyxl), .xls (xlrd)
- CSV/TXT: Auto-detects common encodings
- Protected/encrypted Excel files will fail to read

**Column Mapping Behavior:**
- Fuzzy matching disabled by default (can cause false positives)
- Minimum keyword length for fuzzy match: 2 characters
- Longer keywords matched first to avoid short-word ambiguity

## Testing Notes

**No automated tests exist.** To verify changes:

1. Run the application: `./run_mac_linux.sh`
2. Add sample Excel/CSV files with varied column names
3. Test with/without "统一列名" option
4. Verify column mapping report shows correct transformations
5. Check output file has expected structure
6. Review log files for errors

**Common test scenarios:**
- Mixed .xlsx, .xls, .csv files
- Files with different encodings (UTF-8, GBK)
- Duplicate column names across files
- Empty/partial sheets
- Very large files (>100MB)
- Files with special characters in column names
