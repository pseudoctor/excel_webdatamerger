"""
列名映射配置管理模块
支持用户自定义列名映射规则，并可保存/加载配置
"""
import json
import os
from typing import Dict, List


class ConfigManager:
    """配置管理器 - 处理列名映射规则的保存、加载和管理"""

    DEFAULT_CONFIG_FILE = "column_mappings.json"

    # 默认映射规则
    DEFAULT_MAPPINGS = {
        "商品条码": ["条形码", "条码", "国条码", "barcode", "UPC", "商品编码"],
        "商品名称": ["名称", "品名", "产品名称", "product name"],
        "品牌": ["品牌名称", "brand", "Brand Name"],
        "产品型号": ["型号", "产品规格", "规格", "model", "sku"],
        "含税销售额": ["最终销售金额(销售金额+优惠券金额)", "销售金额", "金额"],
        "数量": ["销售数量", "qty", "quantity"],
        "日期": ["订单日期", "date", "订单时间"],
        "单价": ["价格", "单位价格", "unit price"],
        "供应商": ["供应商名称", "supplier"],
        "客户": ["客户名称", "customer"],
    }

    def __init__(self, config_dir: str = None):
        """
        初始化配置管理器

        Args:
            config_dir: 配置文件目录，默认为项目根目录
        """
        if config_dir is None:
            # 默认配置文件在项目根目录
            config_dir = os.path.dirname(os.path.dirname(__file__))

        self.config_dir = config_dir
        self.config_path = os.path.join(config_dir, self.DEFAULT_CONFIG_FILE)
        self.mappings = self._load_mappings()

    def _load_mappings(self) -> Dict[str, List[str]]:
        """加载映射规则，如果文件不存在则使用默认规则"""
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    # 验证格式
                    if isinstance(loaded, dict):
                        return loaded
            except Exception as e:
                print(f"加载配置文件失败: {e}，使用默认配置")

        # 使用默认配置
        return self.DEFAULT_MAPPINGS.copy()

    def save_mappings(self, mappings: Dict[str, List[str]] = None) -> bool:
        """
        保存映射规则到文件

        Args:
            mappings: 映射规则字典，如果为None则保存当前规则

        Returns:
            是否保存成功
        """
        if mappings is not None:
            self.mappings = mappings

        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.mappings, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False

    def add_mapping(self, standard_name: str, aliases: List[str]) -> None:
        """
        添加或更新一个映射规则

        Args:
            standard_name: 标准列名
            aliases: 别名列表
        """
        self.mappings[standard_name] = aliases

    def remove_mapping(self, standard_name: str) -> bool:
        """
        删除一个映射规则

        Args:
            standard_name: 要删除的标准列名

        Returns:
            是否删除成功
        """
        if standard_name in self.mappings:
            del self.mappings[standard_name]
            return True
        return False

    def get_mappings(self) -> Dict[str, List[str]]:
        """获取当前的映射规则"""
        return self.mappings.copy()

    def reset_to_default(self) -> None:
        """重置为默认映射规则"""
        self.mappings = self.DEFAULT_MAPPINGS.copy()

    def get_all_aliases(self) -> List[str]:
        """获取所有别名（扁平化列表）"""
        aliases = []
        for alias_list in self.mappings.values():
            aliases.extend(alias_list)
        return aliases

    def find_standard_name(self, column_name: str) -> str:
        """
        根据列名查找对应的标准列名

        Args:
            column_name: 原始列名

        Returns:
            标准列名，如果没有匹配则返回原列名
        """
        column_lower = column_name.lower().strip()

        # 先精确匹配
        for standard, aliases in self.mappings.items():
            if column_lower == standard.lower():
                return standard
            for alias in aliases:
                if column_lower == alias.lower():
                    return standard

        # 如果没有精确匹配，返回原列名
        return column_name

    def export_template(self, output_path: str) -> bool:
        """
        导出映射规则模板文件

        Args:
            output_path: 输出文件路径

        Returns:
            是否导出成功
        """
        template = {
            "_说明": "这是列名映射配置文件，格式为 '标准列名': ['别名1', '别名2', ...]",
            "_示例": {
                "商品条码": ["条形码", "条码", "barcode"]
            },
            **self.DEFAULT_MAPPINGS
        }

        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"导出模板失败: {e}")
            return False
