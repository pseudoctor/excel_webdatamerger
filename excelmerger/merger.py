"""
核心数据处理模块 - 增强版
支持智能列名归一化、数据验证和统计
"""
import re
from typing import Dict, List, Tuple

import pandas as pd

from .config_manager import ConfigManager


# ================================================
# 工具函数
# ================================================
def normalize_text(text: str) -> str:
    """
    列名清洗标准化：去除空格、符号、统一大小写
    注意：保留 + 号，因为它可能是列名的一部分

    Args:
        text: 原始文本

    Returns:
        标准化后的文本
    """
    if pd.isna(text):
        return ""
    text = str(text).strip()
    # 仅去除换行符和多余空格，保留其他字符用于精确匹配
    text = re.sub(r"[\n\r]+", "", text)
    text = re.sub(r"\s+", "", text)  # 去除所有空格
    return text.lower()  # 统一小写


# ================================================
# 主类
# ================================================
class ExcelMergerCore:
    """核心数据清洗与列名归一化逻辑 - 增强版"""

    def __init__(self, config_manager: ConfigManager = None):
        """
        初始化合并核心

        Args:
            config_manager: 配置管理器，如果为None则创建默认配置
        """
        self.config_manager = config_manager or ConfigManager()
        self.alias_map = self._build_alias_map()
        self.mapping_report = {}  # 存储列名映射报告

    def _build_alias_map(self) -> Dict[str, str]:
        """
        构建列名映射字典
        格式: {normalized_alias: standard_name}

        Returns:
            映射字典
        """
        alias_map = {}
        mappings = self.config_manager.get_mappings()

        for standard, aliases in mappings.items():
            # 标准名自己映射自己
            norm_std = normalize_text(standard)
            alias_map[norm_std] = standard

            # 每个别名映射到标准名
            for alias in aliases:
                norm_alias = normalize_text(alias)
                if norm_alias:  # 忽略空字符串
                    alias_map[norm_alias] = standard

        return alias_map

    def reload_config(self) -> None:
        """重新加载配置（当配置被修改后调用）"""
        self.alias_map = self._build_alias_map()

    def normalize_columns(self, df: pd.DataFrame, enable_fuzzy: bool = False) -> pd.DataFrame:
        """
        统一相似列名（精确匹配 + 可选模糊匹配）
        注意：如果映射后出现重复列名，会自动添加后缀

        Args:
            df: 数据框
            enable_fuzzy: 是否启用模糊匹配（默认关闭，避免误判）

        Returns:
            列名归一化后的数据框
        """
        new_cols = []
        self.mapping_report = {}  # 重置报告

        for col in df.columns:
            norm_col = normalize_text(col)
            matched_name = None

            # 1. 精确匹配（优先级最高）
            if norm_col in self.alias_map:
                matched_name = self.alias_map[norm_col]
                self.mapping_report[col] = (matched_name, "精确匹配")
                new_cols.append(matched_name)
                continue

            # 2. 模糊匹配（可选，按关键词长度倒序匹配）
            if enable_fuzzy:
                matched_name = self._fuzzy_match(norm_col)
                if matched_name:
                    self.mapping_report[col] = (matched_name, "模糊匹配")
                    new_cols.append(matched_name)
                    continue

            # 3. 无匹配，保留原列名
            self.mapping_report[col] = (col, "未映射")
            new_cols.append(col)

        # 检查并处理重复列名
        new_cols = self._ensure_unique_columns(new_cols)
        df.columns = new_cols
        return df

    def _ensure_unique_columns(self, columns: List[str]) -> List[str]:
        """
        确保列名唯一，如果有重复则添加后缀 _1, _2, ...

        Args:
            columns: 列名列表

        Returns:
            唯一的列名列表
        """
        seen = {}
        result = []

        for col in columns:
            if col not in seen:
                seen[col] = 0
                result.append(col)
            else:
                seen[col] += 1
                result.append(f"{col}_{seen[col]}")

        return result

    def _fuzzy_match(self, norm_col: str) -> str:
        """
        模糊匹配：按关键词长度倒序匹配，避免短词误判

        Args:
            norm_col: 标准化后的列名

        Returns:
            匹配到的标准列名，如果没有匹配则返回空字符串
        """
        # 按别名长度倒序排序（长的优先匹配）
        sorted_aliases = sorted(
            self.alias_map.items(),
            key=lambda x: len(x[0]),
            reverse=True
        )

        for alias_key, std in sorted_aliases:
            if alias_key and len(alias_key) >= 2:  # 至少2个字符才进行模糊匹配
                if alias_key in norm_col:
                    return std

        return ""

    def get_mapping_report(self) -> Dict[str, Tuple[str, str]]:
        """
        获取最近一次列名映射的报告

        Returns:
            字典: {原列名: (映射后列名, 匹配类型)}
        """
        return self.mapping_report.copy()

    def validate_data(self, df: pd.DataFrame) -> Dict:
        """
        验证数据质量

        Args:
            df: 数据框

        Returns:
            验证报告字典
        """
        report = {
            "总行数": len(df),
            "总列数": len(df.columns),
            "空值统计": {},
            "重复行数": 0,
            "数据类型": {}
        }

        # 空值统计
        for col in df.columns:
            null_count = df[col].isna().sum()
            null_percent = (null_count / len(df) * 100) if len(df) > 0 else 0
            report["空值统计"][col] = {
                "数量": int(null_count),
                "百分比": round(null_percent, 2)
            }

        # 重复行统计
        report["重复行数"] = int(df.duplicated().sum())

        # 数据类型
        for col in df.columns:
            report["数据类型"][col] = str(df[col].dtype)

        return report

    def deduplicate_smart(
        self,
        df: pd.DataFrame,
        key_columns: List[str] = None,
        keep: str = 'first'
    ) -> pd.DataFrame:
        """
        智能去重

        Args:
            df: 数据框
            key_columns: 关键字段列表，如果为None则对所有列去重
            keep: 保留策略 ('first', 'last', False)

        Returns:
            去重后的数据框
        """
        if key_columns:
            # 检查关键字段是否存在
            existing_keys = [k for k in key_columns if k in df.columns]
            if existing_keys:
                return df.drop_duplicates(subset=existing_keys, keep=keep)

        # 默认全行去重
        return df.drop_duplicates(keep=keep)

    def get_summary_stats(self, df: pd.DataFrame) -> str:
        """
        获取数据汇总统计（用于日志输出）

        Args:
            df: 数据框

        Returns:
            统计摘要字符串
        """
        stats = []
        stats.append(f"数据行数: {len(df)}")
        stats.append(f"数据列数: {len(df.columns)}")

        # 数值列统计
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            stats.append(f"数值列数: {len(numeric_cols)}")

        # 空值率
        total_cells = len(df) * len(df.columns)
        null_cells = df.isna().sum().sum()
        null_rate = (null_cells / total_cells * 100) if total_cells > 0 else 0
        stats.append(f"空值率: {null_rate:.2f}%")

        return " | ".join(stats)

