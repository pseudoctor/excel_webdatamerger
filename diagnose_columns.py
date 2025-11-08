#!/usr/bin/env python3
"""
列名诊断工具 - 帮助查看原始文件中的列名
"""
import sys
from excelmerger.io_utils import read_file

def diagnose_file(file_path):
    """诊断文件的列名"""
    print(f"\n{'='*60}")
    print(f"文件: {file_path}")
    print(f"{'='*60}")

    try:
        sheets = read_file(file_path)
        for sheet_name, df in sheets.items():
            print(f"\n工作表: {sheet_name}")
            print(f"列数: {len(df.columns)}")
            print(f"\n原始列名:")
            for i, col in enumerate(df.columns, 1):
                # 显示列名的每个字符，方便识别相似字符
                char_breakdown = ' '.join([f"{c}(U+{ord(c):04X})" for c in str(col)[:20]])
                print(f"  {i}. {col}")
                if len(str(col)) <= 20:
                    print(f"     字符: {char_breakdown}")
    except Exception as e:
        print(f"❌ 读取失败: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python3 diagnose_columns.py <文件路径>")
        print("示例: python3 diagnose_columns.py data.xlsx")
        sys.exit(1)

    for file_path in sys.argv[1:]:
        diagnose_file(file_path)

    print(f"\n{'='*60}")
    print("提示: 检查上述列名是否与 column_mappings.json 中的配置匹配")
    print("特别注意相似的汉字，如 '含' 和 '合'、'销' 和其他字")
    print(f"{'='*60}\n")
