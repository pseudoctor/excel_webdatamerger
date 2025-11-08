import os

import pandas as pd


def read_file(file_path):
    """
    智能读取 Excel / CSV / TXT 文件。
    自动识别文件类型、编码和读取引擎。
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext in [".xlsx", ".xls"]:
        engines = ["openpyxl", "xlrd"]
        for engine in engines:
            try:
                return pd.read_excel(file_path, sheet_name=None, engine=engine)
            except Exception as e:
                if "File is not a zip file" in str(e) or "not supported" in str(e):
                    continue
                else:
                    raise RuntimeError(f"Excel 文件读取失败: {file_path} ({e})")
        raise RuntimeError(f"Excel 文件无法识别，请确认格式正确: {file_path}")

    elif ext in [".csv", ".txt"]:
        encodings = ["utf-8-sig", "utf-8", "gbk", "latin1"]
        for enc in encodings:
            try:
                df = pd.read_csv(file_path, sep=None, engine="python", encoding=enc)
                return {os.path.basename(file_path): df}
            except Exception:
                continue
        raise RuntimeError(f"无法识别 CSV/TXT 文件编码: {file_path}")

    else:
        raise RuntimeError(f"不支持的文件类型: {file_path}")

def save_to_excel(df, output_path):
    """
    保存结果为 Excel 文件

    Args:
        df: 数据框
        output_path: 输出文件路径
    """
    # 修复：只有当目录路径非空时才创建目录
    output_dir = os.path.dirname(output_path)
    if output_dir:  # 避免空字符串导致的错误
        os.makedirs(output_dir, exist_ok=True)

    df.to_excel(output_path, index=False, engine="openpyxl")

def save_file(df, output_path, file_format="xlsx"):
    """
    保存结果为指定格式的文件（支持 Excel 和 CSV）

    Args:
        df: 数据框
        output_path: 输出文件路径
        file_format: 文件格式，'xlsx' 或 'csv'
    """
    # 修复：只有当目录路径非空时才创建目录
    output_dir = os.path.dirname(output_path)
    if output_dir:  # 避免空字符串导致的错误
        os.makedirs(output_dir, exist_ok=True)

    if file_format.lower() == "csv":
        # 保存为CSV格式，使用UTF-8编码（带BOM以便Excel正确打开）
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
    else:
        # 默认保存为Excel格式
        df.to_excel(output_path, index=False, engine="openpyxl")
