#!/usr/bin/env python3
"""
高级列名诊断工具 - 详细分析列名差异和映射情况
"""
import sys
from collections import defaultdict
from excelmerger.io_utils import read_file
from excelmerger.merger import normalize_text, ExcelMergerCore
from excelmerger.config_manager import ConfigManager

def show_char_details(text):
    """显示字符串的详细信息"""
    chars = []
    for i, char in enumerate(str(text)):
        code = ord(char)
        # 判断是否为不可见字符
        if code < 0x20 or code == 0x7F:
            char_type = "控制字符"
        elif code in [0x3000, 0x00A0, 0x2003, 0x2002, 0x2009]:
            char_type = "空格类"
        elif 0xFF01 <= code <= 0xFF5E:
            char_type = "全角"
        elif char.isspace():
            char_type = "空白"
        else:
            char_type = "正常"

        chars.append({
            'pos': i,
            'char': char if char.isprintable() else f'\\u{code:04x}',
            'code': f'U+{code:04X}',
            'type': char_type
        })
    return chars

def diagnose_files(*file_paths):
    """诊断多个文件的列名映射情况"""

    print("=" * 80)
    print("高级列名诊断工具")
    print("=" * 80)

    # 初始化配置管理器和合并核心
    config_manager = ConfigManager()
    merger = ExcelMergerCore(config_manager)

    all_columns = defaultdict(list)  # 收集所有列名
    normalized_map = defaultdict(list)  # 标准化后的映射

    # 第一步：读取所有文件的列名
    print("\n第一步：读取所有文件的列名")
    print("-" * 80)

    for file_path in file_paths:
        try:
            sheets = read_file(file_path)
            for sheet_name, df in sheets.items():
                source = f"{file_path}::{sheet_name}"
                print(f"\n文件: {file_path}")
                print(f"工作表: {sheet_name}")
                print(f"列数: {len(df.columns)}")
                print()

                for idx, col in enumerate(df.columns, 1):
                    col_str = str(col)
                    normalized = normalize_text(col_str)

                    # 收集信息
                    all_columns[col_str].append(source)
                    normalized_map[normalized].append((col_str, source))

                    # 显示列名详情
                    print(f"  列 {idx}: '{col_str}'")
                    print(f"  标准化: '{normalized}'")

                    # 显示字符详情
                    if len(col_str) <= 30:  # 只显示较短的列名详情
                        char_details = show_char_details(col_str)
                        if any(c['type'] != '正常' for c in char_details):
                            print("  字符详情:")
                            for c in char_details:
                                print(f"    位置{c['pos']}: '{c['char']}' {c['code']} ({c['type']})")
                    print()

        except Exception as e:
            print(f"❌ 读取失败: {file_path} - {e}\n")

    # 第二步：检查标准化冲突
    print("\n" + "=" * 80)
    print("第二步：检查标准化后的列名冲突")
    print("-" * 80)

    conflicts = []
    for normalized, originals in normalized_map.items():
        if len(originals) > 1:
            # 检查是否是真正的冲突（不同的原始列名映射到同一个标准名）
            unique_originals = set(orig for orig, _ in originals)
            if len(unique_originals) > 1:
                conflicts.append((normalized, originals))

    if conflicts:
        print("\n⚠️  发现冲突：多个不同的原始列名被标准化为同一个名称")
        print()
        for normalized, originals in conflicts:
            print(f"标准化后: '{normalized}'")
            unique_originals = {}
            for orig, source in originals:
                if orig not in unique_originals:
                    unique_originals[orig] = []
                unique_originals[orig].append(source)

            for orig, sources in unique_originals.items():
                print(f"  ← 原始列名: '{orig}'")
                for src in sources:
                    print(f"     来源: {src}")
                # 显示两个列名的差异
                if len(unique_originals) == 2:
                    orig_list = list(unique_originals.keys())
                    if orig != orig_list[0]:
                        compare_strings(orig_list[0], orig_list[1])
            print()
    else:
        print("✅ 未发现标准化冲突")

    # 第三步：模拟列名映射
    print("\n" + "=" * 80)
    print("第三步：模拟列名映射过程")
    print("-" * 80)

    print("\n当前映射配置:")
    mappings = config_manager.get_mappings()
    for std_name, aliases in list(mappings.items())[:5]:  # 只显示前5个
        print(f"  {std_name}: {aliases}")
    print(f"  ... (共 {len(mappings)} 个映射规则)")

    print("\n模拟映射结果:")
    mapped_results = defaultdict(list)

    for normalized, originals in normalized_map.items():
        unique_originals = set(orig for orig, _ in originals)
        for orig in unique_originals:
            # 检查是否在映射表中
            if normalized in merger.alias_map:
                mapped_name = merger.alias_map[normalized]
                mapped_results[mapped_name].append(orig)
                print(f"  '{orig}' → '{mapped_name}' ✓")
            else:
                mapped_results[orig].append(orig)
                print(f"  '{orig}' → '{orig}' (未映射)")

    # 第四步：检查映射后的冲突
    print("\n" + "=" * 80)
    print("第四步：检查映射后是否会产生重复列名")
    print("-" * 80)

    final_conflicts = []
    for mapped_name, originals in mapped_results.items():
        if len(set(originals)) > 1:
            final_conflicts.append((mapped_name, originals))

    if final_conflicts:
        print("\n❌ 发现问题：以下列名在映射后会产生冲突")
        print("   这会导致程序自动添加 _1, _2 等后缀")
        print()
        for mapped_name, originals in final_conflicts:
            print(f"映射后的列名: '{mapped_name}'")
            print(f"来自以下原始列名:")
            for orig in set(originals):
                print(f"  - '{orig}'")
            print()
    else:
        print("\n✅ 未发现映射后的冲突")

    print("=" * 80)

def compare_strings(s1, s2):
    """比较两个字符串的差异"""
    print(f"\n  详细比较:")
    print(f"    字符串1: '{s1}' (长度: {len(s1)})")
    print(f"    字符串2: '{s2}' (长度: {len(s2)})")

    if len(s1) != len(s2):
        print(f"    ⚠️ 长度不同！")

    # 逐字符比较
    max_len = max(len(s1), len(s2))
    for i in range(max_len):
        c1 = s1[i] if i < len(s1) else '(无)'
        c2 = s2[i] if i < len(s2) else '(无)'

        if c1 != c2:
            code1 = f'U+{ord(c1):04X}' if c1 != '(无)' else 'N/A'
            code2 = f'U+{ord(c2):04X}' if c2 != '(无)' else 'N/A'
            print(f"    位置{i}: '{c1}' {code1} ≠ '{c2}' {code2} ⚠️")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python3 advanced_diagnose.py <文件路径1> [文件路径2] ...")
        print("示例: python3 advanced_diagnose.py file1.xlsx file2.csv")
        sys.exit(1)

    diagnose_files(*sys.argv[1:])
