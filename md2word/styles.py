"""docx 样式应用：页面设置、中/英文字体、行距、对齐。"""

from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

PAGE_SIZES_CM = {
    "A4": (21.0, 29.7),
    "A5": (14.8, 21.0),
    "LETTER": (21.59, 27.94),
}

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def apply_page_setup(doc, page_cfg: dict) -> None:
    """设置页面大小与四边页边距。"""
    section = doc.sections[0]
    size = str(page_cfg.get("size", "A4"))
    if size.upper() in PAGE_SIZES_CM:
        width_cm, height_cm = PAGE_SIZES_CM[size.upper()]
    else:  # custom
        width_cm = float(page_cfg["width_cm"])
        height_cm = float(page_cfg["height_cm"])
    section.page_width = Cm(width_cm)
    section.page_height = Cm(height_cm)
    section.top_margin = Cm(float(page_cfg["margin_top_cm"]))
    section.bottom_margin = Cm(float(page_cfg["margin_bottom_cm"]))
    section.left_margin = Cm(float(page_cfg["margin_left_cm"]))
    section.right_margin = Cm(float(page_cfg["margin_right_cm"]))


def set_run_font(run, font_en=None, font_zh=None, font_size=None,
                 bold=None, italic=None, color=None, underline=None) -> None:
    """设置 run 字体：font_en 作用于西文(ascii/hAnsi)，font_zh 作用于中文(eastAsia)。"""
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    if font_en:
        rfonts.set(qn("w:ascii"), font_en)
        rfonts.set(qn("w:hAnsi"), font_en)
    if font_zh:
        rfonts.set(qn("w:eastAsia"), font_zh)
    if font_size is not None:
        run.font.size = Pt(float(font_size))
    if bold is not None:
        run.font.bold = bool(bold)
    if italic is not None:
        run.font.italic = bool(italic)
    if underline is not None:
        run.font.underline = bool(underline)
    if color:
        run.font.color.rgb = RGBColor.from_string(str(color).lstrip("#"))


def style_run(run, font_cfg: dict, **overrides) -> None:
    """按配置块（body/headings.hN/code）给 run 应用字体样式。"""
    set_run_font(
        run,
        font_en=overrides.pop("font_en", font_cfg.get("font_en")),
        font_zh=overrides.pop("font_zh", font_cfg.get("font_zh")),
        font_size=overrides.pop("font_size", font_cfg.get("font_size")),
        bold=overrides.pop("bold", font_cfg.get("bold")),
        italic=overrides.pop("italic", font_cfg.get("italic")),
        color=overrides.pop("color", font_cfg.get("color")),
        underline=overrides.pop("underline", font_cfg.get("underline")),
    )


def apply_paragraph_format(paragraph, *, align=None, line_spacing=None,
                           space_before=None, space_after=None,
                           left_indent_cm=None, first_line_indent_chars=0,
                           font_size=None) -> None:
    """设置段落格式：对齐/行距/段前段后间距/缩进。"""
    pf = paragraph.paragraph_format
    if align:
        paragraph.alignment = ALIGN_MAP[str(align).lower()]
    if line_spacing is not None:
        pf.line_spacing = float(line_spacing)
    if space_before is not None:
        pf.space_before = Pt(float(space_before))
    if space_after is not None:
        pf.space_after = Pt(float(space_after))
    if left_indent_cm is not None:
        pf.left_indent = Cm(float(left_indent_cm))
    if first_line_indent_chars and font_size:
        # 首行缩进按字符宽度折算，避免依赖 Word 的字符单位设置
        pf.first_line_indent = Pt(float(font_size) * first_line_indent_chars)
