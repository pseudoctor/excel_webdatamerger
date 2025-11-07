# 更新日志

## v1.0 (2025-11-06)

### 初始发布

**excel_datamerger** - Excel/CSV/TXT 数据合并工具

### 核心功能

- 多格式支持：Excel (.xlsx, .xls), CSV (.csv), TXT (.txt)
- 智能列名映射：用户自定义映射规则
- 数据质量检查：空值统计、重复检测
- 智能去重：支持基于关键字段去重
- 实时日志：详细的操作记录
- 跨平台支持：Windows, macOS, Linux

### 技术特性

- 精确匹配优先的列名归一化算法
- 自动处理重复列名（添加后缀）
- 多编码自动检测（UTF-8, GBK, Latin1）
- 线程化处理避免GUI阻塞
- 深色模式界面优化

### 使用方式

**Windows:**
```cmd
双击 run_windows.bat
```

**macOS/Linux:**
```bash
./run_mac_linux.sh
```

---

**许可证:** MIT License
