# Windows 故障排除指南

## 问题：ModuleNotFoundError: No module named 'openpyxl'

这个错误说明虚拟环境中的依赖没有正确安装。

## 解决方案

### 方案 1：使用修复脚本（推荐）

运行虚拟环境修复工具：
```cmd
fix_venv.bat
```

这个脚本会：
1. 删除旧的虚拟环境
2. 创建新的虚拟环境
3. 安装所有依赖
4. 验证安装是否成功

### 方案 2：使用检查工具

先运行环境检查工具诊断问题：
```cmd
check_env.bat
```

这个脚本会检查：
- Python 是否安装
- 虚拟环境是否存在
- 依赖文件是否存在
- 已安装的包
- 模块是否能正常导入

根据检查结果，你可以决定是否需要运行 `fix_venv.bat` 修复。

### 方案 3：手动修复

如果自动脚本无法解决问题，可以手动执行以下步骤：

1. **删除旧的虚拟环境**
   ```cmd
   rmdir /s /q venv
   ```

2. **创建新的虚拟环境**
   ```cmd
   python -m venv venv
   ```
   或
   ```cmd
   py -m venv venv
   ```

3. **激活虚拟环境**
   ```cmd
   venv\Scripts\activate
   ```

4. **升级 pip**
   ```cmd
   python -m pip install --upgrade pip
   ```

5. **安装依赖**
   ```cmd
   pip install -r requirements.txt
   ```

6. **验证安装**
   ```cmd
   python -c "import pandas, openpyxl, xlrd, chardet; print('所有依赖已安装')"
   ```

7. **运行程序**
   ```cmd
   python main.py
   ```

## 常见问题

### Q: 执行 fix_venv.bat 时提示 "拒绝访问"
**A:** 以管理员身份运行命令提示符，或者手动删除 venv 文件夹后重试。

### Q: pip install 速度很慢或失败
**A:** 可以使用国内镜像源加速：
```cmd
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### Q: Python 版本不兼容
**A:** 本程序需要 Python 3.9 或更高版本。检查版本：
```cmd
python --version
```
如果版本过低，请从 https://www.python.org/downloads/ 下载最新版本。

### Q: 虚拟环境激活后仍然找不到模块
**A:** 确保你在激活虚拟环境后使用的是 `python` 命令而不是 `py` 命令：
```cmd
venv\Scripts\activate
python main.py    # ✅ 正确
py main.py        # ❌ 错误（可能使用系统 Python）
```

## 工具脚本说明

### run_windows.bat
- **功能**：启动程序的主脚本
- **改进**：增加了详细的错误日志和依赖验证
- **用法**：双击运行或在命令行执行

### fix_venv.bat
- **功能**：重建虚拟环境
- **用途**：当依赖安装出错时使用
- **特点**：会删除旧环境并重新创建

### check_env.bat
- **功能**：诊断环境问题
- **用途**：检查 Python、虚拟环境、依赖状态
- **特点**：不会修改任何文件，仅进行检查

## 预防措施

为了避免将来出现类似问题：

1. **不要手动修改 venv 文件夹**
2. **使用 requirements.txt 管理依赖**
3. **定期清理并重建虚拟环境**（尤其是更新 Python 版本后）
4. **保持网络连接稳定**（安装依赖时）

## 联系支持

如果以上方法都无法解决问题，请提供以下信息：

1. `check_env.bat` 的完整输出
2. Python 版本 (`python --version`)
3. Windows 版本
4. 错误信息截图

---

*最后更新：2025-11-13*
