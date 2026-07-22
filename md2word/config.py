"""配置加载：默认配置 + YAML 用户配置深合并。"""

from __future__ import annotations

import copy

import yaml

DEFAULT_CONFIG: dict = {
    "page": {
        "size": "A4",  # A4 / A5 / Letter / custom（custom 时用 width_cm/height_cm）
        "width_cm": 21.0,
        "height_cm": 29.7,
        "margin_top_cm": 2.54,
        "margin_bottom_cm": 2.54,
        "margin_left_cm": 3.17,
        "margin_right_cm": 3.17,
    },
    "body": {
        "font_en": "Times New Roman",
        "font_zh": "宋体",
        "font_size": 12,          # 磅
        "line_spacing": 1.5,      # 倍数行距
        "first_line_indent": False,  # 正文段落首行缩进两字符
        "color": "000000",        # 默认黑色
    },
    # h1~h6：每级标题可配置 font_en/font_zh/font_size/bold/italic/color/align/
    # space_before/space_after；缺省级别按下面的默认值递减
    "headings": {
        "h1": {"font_en": "Arial", "font_zh": "黑体", "font_size": 22, "bold": True,
               "color": "000000", "align": "center",
               "space_before_lines": 2.0, "space_after_lines": 1.5, "line_spacing": 1.5,
               "space_before": 24, "space_after": 18},
        "h2": {"font_en": "Arial", "font_zh": "黑体", "font_size": 18, "bold": True,
               "color": "000000",
               "space_before_lines": 1.5, "space_after_lines": 1.0, "line_spacing": 1.5,
               "space_before": 18, "space_after": 12},
        "h3": {"font_en": "Arial", "font_zh": "黑体", "font_size": 16, "bold": True,
               "color": "000000",
               "space_before_lines": 1.0, "space_after_lines": 0.5, "line_spacing": 1.5,
               "space_before": 13, "space_after": 6},
        "h4": {"font_en": "Arial", "font_zh": "黑体", "font_size": 14, "bold": True,
               "color": "000000",
               "space_before_lines": 0.5, "space_after_lines": 0.5, "line_spacing": 1.5,
               "space_before": 12, "space_after": 6},
        "h5": {"font_en": "Arial", "font_zh": "黑体", "font_size": 12, "bold": True,
               "color": "000000",
               "space_before_lines": 0.5, "space_after_lines": 0.5, "line_spacing": 1.5,
               "space_before": 12, "space_after": 6},
        "h6": {"font_en": "Arial", "font_zh": "黑体", "font_size": 11, "bold": True,
               "color": "000000",
               "space_before_lines": 0.5, "space_after_lines": 0.5, "line_spacing": 1.5,
               "space_before": 6, "space_after": 6},
    },
    "code": {
        "font_en": "Consolas",
        "font_zh": "宋体",
        "font_size": 10.5,
        "line_spacing": 1.15,
        "color": "000000",
    },
    "quote": {
        "italic": False,
        "indent_cm": 0.75,
    },
    "table": {
        "style": "grid",          # grid（全边框）/ three_line（三线表）
        "header_align": "center",  # 表头对齐：left / center / right
        "header_bold": True,
    },
}


def deep_merge(base: dict, override: dict) -> dict:
    """把 override 递归合并进 base（返回新 dict）。"""
    result = copy.deepcopy(base)
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def load_config(path: str | None = None) -> dict:
    """加载配置：无路径时返回默认配置；否则与默认配置深合并。"""
    if not path:
        return copy.deepcopy(DEFAULT_CONFIG)
    with open(path, "r", encoding="utf-8") as f:
        user_cfg = yaml.safe_load(f) or {}
    if not isinstance(user_cfg, dict):
        raise ValueError(f"配置文件格式错误：{path}（顶层必须是键值对）")
    cfg = deep_merge(DEFAULT_CONFIG, user_cfg)
    # 标题级别键规整：允许 "1"/1 写法，统一为 h1..h6
    headings = cfg.get("headings", {})
    for level in range(1, 7):
        for alias in (str(level), level):
            if alias in headings:
                headings[f"h{level}"] = headings.pop(alias)
    return cfg
