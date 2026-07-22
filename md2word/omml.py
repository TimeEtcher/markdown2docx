"""LaTeX 公式 -> Word OMML（Office Math ML）转换。

链路：LaTeX --(latex2mathml)--> MathML --(本模块)--> OMML，
生成的 <m:oMath> 元素可直接挂到 docx 的 <w:p> 下，Word 中显示为原生可编辑公式。
"""

from __future__ import annotations

import latex2mathml.converter
from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# n-ary 运算符（上下限放正上/正下方，而不是右下角）
NARY_OPS = set("∑∏∐∫∬∭∮⋂⋃⋀⋁")

# 可视为重音符号的字符（\hat/\bar/\vec/\tilde 等）
ACCENT_CHARS = {"^", "¯", "~", "→", "←", "˙", "¨", "ˇ", "˘", "˚"}

# 围栏字符映射（mo -> 显示用字符）
FENCE_DISPLAY = {
    "(": "(", ")": ")",
    "[": "[", "]": "]",
    "{": "{", "}": "}",
    "|": "|", "‖": "‖",
    "⟨": "⟨", "⟩": "⟩",
    "⌈": "⌈", "⌉": "⌉",
    "⌊": "⌊", "⌋": "⌋",
}


class MathConversionError(Exception):
    pass


def _m(parent, tag):
    return etree.SubElement(parent, f"{{{M}}}{tag}")


def _local(node) -> str:
    return etree.QName(node).localname


def latex_to_omml(latex: str, font_size: float | None = None) -> etree._Element:
    """返回 <m:oMath> 元素；font_size（磅）用于让公式字号与正文一致。"""
    try:
        mathml = latex2mathml.converter.convert(latex)
        root = etree.fromstring(mathml.encode("utf-8"))
    except Exception as exc:
        raise MathConversionError(str(exc)) from exc
    omath = etree.Element(f"{{{M}}}oMath", nsmap={"m": M, "w": W})
    _convert_children(root, omath, font_size)
    return omath


def wrap_omath_para(omath: etree._Element) -> etree._Element:
    """把 <m:oMath> 包一层 <m:oMathPara>（独立成段的显示公式用）。"""
    para = etree.Element(f"{{{M}}}oMathPara", nsmap={"m": M, "w": W})
    para.append(omath)
    return para


# ------------------------------------------------------------ MathML -> OMML
def _convert_children(node, parent, size):
    children = list(node)
    i = 0
    while i < len(children):
        child = children[i]
        tag = _local(child)

        # 左右定界符：mo fence=prefix ... mo fence=postfix -> m:d
        if tag == "mo" and child.get("fence") == "true" and child.get("form") == "prefix":
            j = _find_matching_fence(children, i)
            if j is not None:
                d = _m(parent, "d")
                pr = _m(d, "dPr")
                beg = _m(pr, "begChr")
                beg.set(f"{{{M}}}val", _fence_display(child.text))
                end = _m(pr, "endChr")
                end.set(f"{{{M}}}val", _fence_display(children[j].text))
                e = _m(d, "e")
                for k in range(i + 1, j):
                    _convert_node(children[k], e, size)
                i = j + 1
                continue

        # n-ary 运算符：sum/integral 等，把后续同级子节点都作为被作用表达式放入 m:e
        if tag in ("msub", "msup", "msubsup") and len(child) > 0 and _is_nary_base(child[0]):
            _convert_node(child, parent, size)
            # _convert_node 刚把 nary 追加到 parent 末尾
            nary = parent[-1]
            if _local(nary) == "nary":
                e = nary.find(f"{{{M}}}e")
                if e is not None:
                    j = i + 1
                    while j < len(children):
                        _convert_node(children[j], e, size)
                        j += 1
                    i = j
                    continue
        else:
            _convert_node(child, parent, size)
        i += 1


def _find_matching_fence(children, start):
    depth = 1
    for j in range(start + 1, len(children)):
        c = children[j]
        if _local(c) != "mo" or c.get("fence") != "true":
            continue
        form = c.get("form")
        if form == "prefix":
            depth += 1
        elif form == "postfix":
            depth -= 1
            if depth == 0:
                return j
    return None


def _fence_display(text):
    return FENCE_DISPLAY.get(text, text)


def _convert_node(node, parent, size):
    tag = _local(node)
    handler = _HANDLERS.get(tag)
    if handler is not None:
        handler(node, parent, size)
    elif tag in ("math", "mrow", "mstyle", "semantics", "mpadded", "merror"):
        _convert_children(node, parent, size)
    # 其余未知标签直接忽略，保证不中断转换


def _add_run(parent, text, size, node=None, plain=False):
    r = _m(parent, "r")
    # 字号与正文保持一致
    if size:
        rpr = etree.SubElement(r, f"{{{W}}}rPr")
        for t in ("sz", "szCs"):
            sz = etree.SubElement(rpr, f"{{{W}}}{t}")
            sz.set(f"{{{W}}}val", str(int(round(float(size) * 2))))
    # 正体：数字、运算符、普通文本
    if plain or (node is not None and node.get("mathvariant") == "normal"):
        mrpr = _m(r, "rPr")
        sty = _m(mrpr, "sty")
        sty.set(f"{{{M}}}val", "p")
    t = _m(r, "t")
    t.text = text
    if text != text.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _tok(node, parent, size):  # mi / mn / mo / mtext
    tag = _local(node)
    text = node.text or ""
    # mi 默认斜体；mn/mo/mtext 用正体
    plain = tag in ("mn", "mo", "mtext")
    _add_run(parent, text, size, node, plain=plain)


def _mspace(node, parent, size):
    _add_run(parent, " ", size, plain=True)


def _mfrac(node, parent, size):
    f = _m(parent, "f")
    num = _m(f, "num")
    _convert_node(node[0], num, size)
    den = _m(f, "den")
    _convert_node(node[1], den, size)


def _msup(node, parent, size):
    if _is_nary_base(node[0]):
        _nary(node[0], None, node[1], parent, size)
        return
    s = _m(parent, "sSup")
    e = _m(s, "e")
    _convert_node(node[0], e, size)
    sup = _m(s, "sup")
    _convert_node(node[1], sup, size)


def _msub(node, parent, size):
    if _is_nary_base(node[0]):
        _nary(node[0], node[1], None, parent, size)
        return
    s = _m(parent, "sSub")
    e = _m(s, "e")
    _convert_node(node[0], e, size)
    sub = _m(s, "sub")
    _convert_node(node[1], sub, size)


def _msubsup(node, parent, size):
    if _is_nary_base(node[0]):
        _nary(node[0], node[1], node[2], parent, size)
        return
    s = _m(parent, "sSubSup")
    e = _m(s, "e")
    _convert_node(node[0], e, size)
    sub = _m(s, "sub")
    _convert_node(node[1], sub, size)
    sup = _m(s, "sup")
    _convert_node(node[2], sup, size)


def _is_nary_base(node) -> bool:
    return _local(node) == "mo" and (node.text or "").strip() in NARY_OPS


def _nary(base, under, over, parent, size):
    nary = _m(parent, "nary")
    pr = _m(nary, "naryPr")
    chr_el = _m(pr, "chr")
    chr_el.set(f"{{{M}}}val", (base.text or "").strip())
    sub = _m(nary, "sub")
    if under is not None:
        _convert_node(under, sub, size)
    sup = _m(nary, "sup")
    if over is not None:
        _convert_node(over, sup, size)
    _m(nary, "e")  # 被作用表达式由 _convert_children 中下一个节点补充


def _munder(node, parent, size):
    base, under = node[0], node[1]
    if _is_nary_base(base):
        _nary(base, under, None, parent, size)
        return
    lim = _m(parent, "limLow")
    e = _m(lim, "e")
    _convert_node(base, e, size)
    l = _m(lim, "lim")
    _convert_node(under, l, size)


def _mover(node, parent, size):
    base, over = node[0], node[1]
    # \hat/\bar/\vec/\tilde 等重音符号：用 m:acc 让字符正上方加符号
    over_text = (over.text or "").strip()
    if _local(over) == "mo" and len(over_text) == 1 and (
            node.get("accent") == "true" or over_text in ACCENT_CHARS):
        acc = _m(parent, "acc")
        pr = _m(acc, "accPr")
        chr_el = _m(pr, "chr")
        chr_el.set(f"{{{M}}}val", over_text)
        e = _m(acc, "e")
        _convert_node(base, e, size)
        return
    if _is_nary_base(base):
        _nary(base, None, over, parent, size)
        return
    lim = _m(parent, "limUpp")
    e = _m(lim, "e")
    _convert_node(base, e, size)
    l = _m(lim, "lim")
    _convert_node(over, l, size)


def _munderover(node, parent, size):
    base, under, over = node[0], node[1], node[2]
    if _is_nary_base(base):
        _nary(base, under, over, parent, size)
        return
    low = _m(parent, "limLow")
    upp = _m(low, "e")
    inner = _m(upp, "limUpp")
    e = _m(inner, "e")
    _convert_node(base, e, size)
    l = _m(inner, "lim")
    _convert_node(over, l, size)
    l2 = _m(low, "lim")
    _convert_node(under, l2, size)


def _msqrt(node, parent, size):
    rad = _m(parent, "rad")
    pr = _m(rad, "radPr")
    deg_hide = _m(pr, "degHide")
    deg_hide.set(f"{{{M}}}val", "1")
    _m(rad, "deg")
    e = _m(rad, "e")
    _convert_node(node[0], e, size)


def _mroot(node, parent, size):
    rad = _m(parent, "rad")
    deg = _m(rad, "deg")
    _convert_node(node[1], deg, size)
    e = _m(rad, "e")
    _convert_node(node[0], e, size)


def _mfenced(node, parent, size):
    d = _m(parent, "d")
    pr = _m(d, "dPr")
    beg = _m(pr, "begChr")
    beg.set(f"{{{M}}}val", node.get("open", "("))
    end = _m(pr, "endChr")
    end.set(f"{{{M}}}val", node.get("close", ")"))
    e = _m(d, "e")
    _convert_children(node, e, size)


def _mtable(node, parent, size):
    m = _m(parent, "m")
    for row in node:
        if _local(row) != "mtr":
            continue
        mr = _m(m, "mr")
        for cell in row:
            e = _m(mr, "e")
            _convert_children(cell, e, size)


_HANDLERS = {
    "mi": _tok, "mn": _tok, "mo": _tok, "mtext": _tok,
    "mspace": _mspace,
    "mfrac": _mfrac,
    "msup": _msup, "msub": _msub, "msubsup": _msubsup,
    "munder": _munder, "mover": _mover, "munderover": _munderover,
    "msqrt": _msqrt, "mroot": _mroot,
    "mfenced": _mfenced,
    "mtable": _mtable,
}
