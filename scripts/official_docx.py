#!/usr/bin/env python3
"""Apply and audit the user's Chinese official-document DOCX format."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
from zipfile import BadZipFile

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from lxml.etree import XMLSyntaxError


TITLE_FONT = "方正小标宋简体"
BODY_FONT = "仿宋_GB2312"
LEVEL1_FONT = "黑体"
LEVEL2_FONT = "楷体_GB2312"
PAGE_FONT = "宋体"
TITLE_SIZE_PT = 22
BODY_SIZE_PT = 16
PAGE_SIZE_PT = 14
LINE_SPACING_PT = 29
BODY_FIRST_LINE_CHARS = 200
MAX_TITLE_UNITS = 20.0
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7
MARGINS_CM = {
    "top": 3.7,
    "bottom": 3.5,
    "left": 2.8,
    "right": 2.6,
}
FOOTER_DISTANCE_CM = 2.8
PAGE_SIZE_TWIPS = (11906, 16838)
MARGIN_TWIPS = {
    "top": 2098,
    "bottom": 1984,
    "left": 1587,
    "right": 1474,
    "gutter": 0,
    "footer": 1587,
}
FONT_FILE_HINTS = {
    TITLE_FONT: ("方正小标宋简体", "fzxbs"),
    BODY_FONT: ("仿宋_gb2312", "fangsong_gb2312"),
    LEVEL1_FONT: ("黑体", "simhei"),
    LEVEL2_FONT: ("楷体_gb2312", "kaiti_gb2312"),
    PAGE_FONT: ("simsun", "宋体"),
}

HEADING_FONT_BY_LEVEL = {
    1: LEVEL1_FONT,
    2: LEVEL2_FONT,
    3: BODY_FONT,
    4: BODY_FONT,
}
_CHINESE_NUMERAL = r"[一二三四五六七八九十百千万零〇]+"
_HEADING_PATTERNS = (
    (4, re.compile(r"^（[1-9][0-9]*）")),
    (2, re.compile(rf"^（{_CHINESE_NUMERAL}）")),
    (3, re.compile(r"^[1-9][0-9]*\.(?![0-9])")),
    (1, re.compile(rf"^{_CHINESE_NUMERAL}、")),
)
_HEADING_VARIANT_PATTERNS = (
    (4, re.compile(r"^\([1-9][0-9]*\)"), "四级标题应使用全角括号（1），不能使用半角括号 (1)。"),
    (2, re.compile(rf"^\({_CHINESE_NUMERAL}\)"), "二级标题应使用全角括号（一），不能使用半角括号 (一)。"),
    (3, re.compile(r"^[1-9][0-9]*．(?![0-9])"), "三级标题应使用半角句点 1.，不能使用全角句点 1．。"),
)
_TERMINAL_PUNCTUATION = frozenset("，。！？；：、,.!?;:…")
_TRAILING_CLOSERS = frozenset("\"'”’）)]】》」』")

_PPR_ORDER = {
    qn("w:pStyle"): 10,
    qn("w:keepNext"): 20,
    qn("w:keepLines"): 30,
    qn("w:spacing"): 100,
    qn("w:ind"): 110,
    qn("w:jc"): 200,
    qn("w:sectPr"): 1000,
}


for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


@dataclass
class AuditReport:
    path: str
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors

    def as_dict(self) -> dict[str, object]:
        return {
            "path": self.path,
            "ok": self.ok,
            "errors": self.errors,
            "warnings": self.warnings,
        }


@dataclass(frozen=True)
class HeadingMatch:
    level: int
    content: str
    leading_whitespace: bool


_INDENT_ATTRIBUTES = ("w:firstLineChars", "w:firstLine", "w:hanging", "w:hangingChars")


def _get_or_add(parent, tag: str):
    child = parent.find(qn(tag))
    if child is None:
        child = OxmlElement(tag)
    if parent.tag != qn("w:pPr"):
        if child.getparent() is None:
            parent.append(child)
        return child

    desired_order = _PPR_ORDER.get(qn(tag), 500)
    if child.getparent() is parent:
        parent.remove(child)
    for existing in list(parent):
        existing_order = _PPR_ORDER.get(existing.tag, 500)
        if existing_order > desired_order:
            parent.insert(parent.index(existing), child)
            break
    else:
        parent.append(child)
    return child


def _remove_all(parent, tag: str) -> None:
    for child in list(parent.findall(qn(tag))):
        parent.remove(child)


def _clear_children(element) -> None:
    for child in list(element):
        element.remove(child)


def _load_document(path: Path):
    try:
        return Document(str(path))
    except (PackageNotFoundError, BadZipFile, OSError, ValueError, KeyError, XMLSyntaxError) as exc:
        raise ValueError(f"无法读取 DOCX 文件：{path}；文件可能损坏或无权访问：{exc}") from exc


def _set_rpr(r_pr, font_name: str, size_pt: int, clean_emphasis: bool = False) -> None:
    r_fonts = _get_or_add(r_pr, "w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        r_fonts.set(qn(attr), font_name)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        expanded = qn(attr)
        if expanded in r_fonts.attrib:
            del r_fonts.attrib[expanded]

    half_points = str(size_pt * 2)
    _get_or_add(r_pr, "w:sz").set(qn("w:val"), half_points)
    _get_or_add(r_pr, "w:szCs").set(qn("w:val"), half_points)

    if clean_emphasis:
        for tag in ("w:b", "w:bCs", "w:i", "w:iCs", "w:u"):
            _remove_all(r_pr, tag)


def _set_run(run_element, font_name: str, size_pt: int, clean_emphasis: bool = False) -> None:
    r_pr = run_element.find(qn("w:rPr"))
    if r_pr is None:
        r_pr = OxmlElement("w:rPr")
        run_element.insert(0, r_pr)
    _set_rpr(r_pr, font_name, size_pt, clean_emphasis)


def _set_paragraph_format(paragraph_element, alignment: str | None = None) -> None:
    p_pr = paragraph_element.find(qn("w:pPr"))
    if p_pr is None:
        p_pr = OxmlElement("w:pPr")
        paragraph_element.insert(0, p_pr)

    spacing = _get_or_add(p_pr, "w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(LINE_SPACING_PT * 20))
    spacing.set(qn("w:lineRule"), "exact")

    if alignment is not None:
        _get_or_add(p_pr, "w:jc").set(qn("w:val"), alignment)


def _set_body_first_line_indent(paragraph_element) -> None:
    p_pr = paragraph_element.find(qn("w:pPr"))
    if p_pr is None:
        p_pr = OxmlElement("w:pPr")
        paragraph_element.insert(0, p_pr)

    indentation = _get_or_add(p_pr, "w:ind")
    for attr in ("w:firstLine", "w:hanging", "w:hangingChars"):
        expanded = qn(attr)
        if expanded in indentation.attrib:
            del indentation.attrib[expanded]
    indentation.set(qn("w:firstLineChars"), str(BODY_FIRST_LINE_CHARS))


def _clear_first_line_and_hanging_indent(paragraph_element) -> None:
    p_pr = paragraph_element.find(qn("w:pPr"))
    if p_pr is None:
        return
    indentation = p_pr.find(qn("w:ind"))
    if indentation is None:
        return
    for attr in ("w:firstLineChars", "w:firstLine", "w:hanging", "w:hangingChars"):
        expanded = qn(attr)
        if expanded in indentation.attrib:
            del indentation.attrib[expanded]
    if not indentation.attrib and len(indentation) == 0:
        p_pr.remove(indentation)


def _set_keep_with_next(paragraph_element) -> None:
    p_pr = paragraph_element.find(qn("w:pPr"))
    if p_pr is None:
        p_pr = OxmlElement("w:pPr")
        paragraph_element.insert(0, p_pr)
    _get_or_add(p_pr, "w:keepNext").set(qn("w:val"), "1")
    _get_or_add(p_pr, "w:keepLines").set(qn("w:val"), "1")


def _set_style(style, font_name: str, size_pt: int, alignment: str | None = None) -> None:
    p_pr = _get_or_add(style.element, "w:pPr")
    r_pr = _get_or_add(style.element, "w:rPr")
    if style.element.index(p_pr) > style.element.index(r_pr):
        style.element.remove(p_pr)
        style.element.insert(style.element.index(r_pr), p_pr)

    _set_rpr(r_pr, font_name, size_pt, clean_emphasis=alignment == "center")
    spacing = _get_or_add(p_pr, "w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(LINE_SPACING_PT * 20))
    spacing.set(qn("w:lineRule"), "exact")
    if alignment is not None:
        _get_or_add(p_pr, "w:jc").set(qn("w:val"), alignment)


def _configure_styles(document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(BODY_SIZE_PT)
    for tag in ("w:rPr", "w:pPr"):
        element = normal.element.find(qn(tag))
        if element is not None:
            _clear_children(element)
    _set_style(normal, BODY_FONT, BODY_SIZE_PT)

    try:
        title_style = document.styles["OfficialTitle"]
    except KeyError:
        title_style = document.styles.add_style("OfficialTitle", WD_STYLE_TYPE.PARAGRAPH)
        title_style.base_style = normal
    title_style.font.name = TITLE_FONT
    title_style.font.size = Pt(TITLE_SIZE_PT)
    for tag in ("w:rPr", "w:pPr"):
        element = title_style.element.find(qn(tag))
        if element is not None:
            _clear_children(element)
    _set_style(title_style, TITLE_FONT, TITLE_SIZE_PT, alignment="center")


def _iter_body_paragraph_elements(document) -> list:
    return list(document.element.body.iter(qn("w:p")))


def _iter_paragraph_run_elements(paragraph_element):
    """Yield runs owned by this paragraph, excluding nested text-box paragraphs."""

    for run_element in paragraph_element.iter(qn("w:r")):
        ancestor = run_element.getparent()
        while ancestor is not None and ancestor is not paragraph_element:
            if ancestor.tag in {qn("w:p"), qn("w:txbxContent")}:
                break
            ancestor = ancestor.getparent()
        if ancestor is paragraph_element:
            yield run_element


def _visible_text(element) -> str:
    parts: list[str] = []
    hidden_containers = {
        qn("w:del"),
        qn("w:moveFrom"),
        qn("w:txbxContent"),
    }

    def visit(node) -> None:
        for child in node:
            if child.tag in hidden_containers or child.tag == qn("w:p"):
                continue
            if child.tag == qn("w:t"):
                parts.append(child.text or "")
            elif child.tag in {qn("w:tab"), qn("w:ptab")}:
                parts.append("\t")
            elif child.tag in {qn("w:br"), qn("w:cr")}:
                parts.append("\n")
            else:
                visit(child)

    visit(element)
    return "".join(parts)


def _classify_heading(paragraph_element) -> HeadingMatch | None:
    raw_text = _visible_text(paragraph_element)
    content = raw_text.lstrip()
    for level, pattern in _HEADING_PATTERNS:
        marker = pattern.match(content)
        if marker:
            return HeadingMatch(
                level=level,
                content=content[marker.end() :],
                leading_whitespace=content != raw_text,
            )
    return None


def _heading_variant_error(paragraph_element) -> str | None:
    content = _visible_text(paragraph_element).lstrip()
    for _, pattern, message in _HEADING_VARIANT_PATTERNS:
        if pattern.match(content):
            return message
    return None


def _ends_with_terminal_punctuation(text: str) -> bool:
    content = text.rstrip()
    while content and content[-1] in _TRAILING_CLOSERS:
        content = content[:-1].rstrip()
    return bool(content and content[-1] in _TERMINAL_PUNCTUATION)


def _heading_punctuation_error(match: HeadingMatch) -> str | None:
    has_punctuation = _ends_with_terminal_punctuation(match.content)
    if match.level == 1 and has_punctuation:
        return "一级标题句末不得有标点符号。"
    if match.level in {3, 4} and not has_punctuation:
        return f"{match.level}级标题句末必须有标点符号。"
    return None


def _heading_has_meaningful_content(match: HeadingMatch) -> bool:
    content = match.content.strip()
    return any(
        not char.isspace() and not unicodedata.category(char).startswith("P")
        for char in content
    )


def _heading_content_errors(document, title_element) -> list[str]:
    errors: list[str] = []
    for index, paragraph_element in enumerate(_iter_top_level_body_paragraph_elements(document)):
        if paragraph_element is title_element:
            continue
        heading = _classify_heading(paragraph_element)
        if heading is None:
            variant_error = _heading_variant_error(paragraph_element)
            if variant_error:
                errors.append(f"正文顶层段落 {index}：{variant_error}")
            continue
        if heading.leading_whitespace:
            errors.append(f"正文顶层段落 {index}：标题编号前不得有空格、Tab 或换行。")
        if not _heading_has_meaningful_content(heading):
            errors.append(f"正文顶层段落 {index}：标题编号后必须有实际标题内容。")
        punctuation_error = _heading_punctuation_error(heading)
        if punctuation_error:
            errors.append(f"正文顶层段落 {index}：{punctuation_error}")
    return errors


def _format_body(document) -> None:
    for paragraph_element in _iter_body_paragraph_elements(document):
        _set_paragraph_format(paragraph_element)
        for run_element in _iter_paragraph_run_elements(paragraph_element):
            _set_run(run_element, BODY_FONT, BODY_SIZE_PT)


def _iter_top_level_body_paragraph_elements(document) -> list:
    """Return body paragraphs outside tables and nested text boxes."""

    body = document.element.body
    excluded_ancestors = {qn("w:tbl"), qn("w:txbxContent")}
    paragraphs = []
    for paragraph_element in body.iter(qn("w:p")):
        ancestor = paragraph_element.getparent()
        excluded = False
        while ancestor is not None and ancestor is not body:
            if ancestor.tag in excluded_ancestors:
                excluded = True
                break
            ancestor = ancestor.getparent()
        if not excluded:
            paragraphs.append(paragraph_element)
    return paragraphs


def _format_top_level_body_paragraphs(document, title_element) -> None:
    normal_style_id = document.styles["Normal"].style_id
    for paragraph_element in _iter_top_level_body_paragraph_elements(document):
        if paragraph_element is title_element:
            continue
        heading = _classify_heading(paragraph_element)
        if heading is None:
            _set_body_first_line_indent(paragraph_element)
            continue

        p_pr = paragraph_element.find(qn("w:pPr"))
        if p_pr is None:
            p_pr = OxmlElement("w:pPr")
            paragraph_element.insert(0, p_pr)
        _get_or_add(p_pr, "w:pStyle").set(qn("w:val"), normal_style_id)
        _remove_all(p_pr, "w:numPr")
        _clear_first_line_and_hanging_indent(paragraph_element)
        _set_keep_with_next(paragraph_element)
        font_name = HEADING_FONT_BY_LEVEL[heading.level]
        for run_element in _iter_paragraph_run_elements(paragraph_element):
            _set_run(run_element, font_name, BODY_SIZE_PT, clean_emphasis=True)


def _title_paragraphs(document) -> list[Paragraph]:
    return [
        Paragraph(element, document._body)
        for element in _iter_top_level_body_paragraph_elements(document)
    ]


def _find_title(document, title_index: int | None, allow_empty: bool = False):
    paragraphs = _title_paragraphs(document)
    if title_index is not None:
        if title_index < 0 or title_index >= len(paragraphs):
            raise ValueError(
                f"--title-index {title_index} 超出范围；正文顶层段落数为 {len(paragraphs)}。"
            )
        paragraph = paragraphs[title_index]
        if not allow_empty and not paragraph.text.strip():
            raise ValueError("指定的标题段落为空。")
        return title_index, paragraph

    for index, paragraph in enumerate(paragraphs):
        if paragraph.text.strip():
            return index, paragraph

    if allow_empty:
        if paragraphs:
            return 0, paragraphs[0]
        paragraph = document.add_paragraph()
        return 0, paragraph
    raise ValueError("文档没有可识别的非空顶层标题段落。")


def _replace_title(paragraph, title_lines: list[str]) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)
    for index, line in enumerate(title_lines):
        if index:
            paragraph.add_run().add_break()
        paragraph.add_run(line)


def _format_title(paragraph) -> None:
    paragraph_properties = paragraph._p.find(qn("w:pPr"))
    if paragraph_properties is not None:
        for child in list(paragraph_properties):
            if child.tag != qn("w:sectPr"):
                paragraph_properties.remove(child)
    paragraph.style = "OfficialTitle"
    _set_paragraph_format(paragraph._p, alignment="center")
    _set_keep_with_next(paragraph._p)
    for run_element in _iter_paragraph_run_elements(paragraph._p):
        run_properties = run_element.find(qn("w:rPr"))
        if run_properties is not None:
            _clear_children(run_properties)
        _set_run(run_element, TITLE_FONT, TITLE_SIZE_PT, clean_emphasis=True)


def _title_units(text: str) -> float:
    units = 0.0
    for char in text.strip():
        width = unicodedata.east_asian_width(char)
        units += 1.0 if width in {"W", "F", "A"} else 0.5
    return units


def _title_shape_errors(lines: Iterable[str]) -> list[str]:
    source_lines = list(lines)
    normalized = [line.strip() for line in source_lines]
    errors: list[str] = []
    if not normalized or any(not line for line in normalized):
        return ["标题行不得为空。"]
    if any("\n" in line or "\r" in line for line in normalized):
        errors.append("每个 --title-line 必须只包含一行。")
    if any(any(char.isspace() for char in line) for line in normalized):
        errors.append("标题不得使用空格或 Tab 造形；请使用显式换行。")
    if any(line != line.lstrip() or line != line.rstrip() for line in source_lines):
        errors.append("标题行不得包含前导或尾随空格。")

    widths = [_title_units(line) for line in normalized]
    too_wide = [
        f"第 {i + 1} 行={width:g}"
        for i, width in enumerate(widths)
        if width > MAX_TITLE_UNITS
    ]
    if too_wide:
        errors.append(
            "标题行超过约 20 个全角视觉单位，可能再次自动换行：" + "，".join(too_wide)
        )

    if len(widths) > 1:
        if max(widths) - min(widths) < 0.5:
            errors.append("多行标题近似等宽，形成矩形；请按语义调整为梯形或菱形。")
        signs: list[int] = []
        for left, right in zip(widths, widths[1:]):
            delta = right - left
            if abs(delta) < 0.25:
                continue
            signs.append(1 if delta > 0 else -1)
        seen_decrease = False
        valley = False
        for sign in signs:
            if sign < 0:
                seen_decrease = True
            elif seen_decrease:
                valley = True
                break
        if valley:
            errors.append("标题行先变短后变长，形成沙漏；请改为梯形或先长后短的菱形。")
    return errors


def _extract_title_lines(paragraph) -> tuple[list[str], list[str]]:
    lines = [""]
    break_types: list[str] = []
    for node in paragraph._p.iter():
        if node.tag in {qn("w:t"), qn("w:delText"), qn("w:instrText")}:
            lines[-1] += node.text or ""
        elif node.tag == qn("w:tab"):
            lines[-1] += "\t"
        elif node.tag == qn("w:br"):
            break_types.append(node.get(qn("w:type")) or "textWrapping")
            lines.append("")
        elif node.tag == qn("w:cr"):
            break_types.append("textWrapping")
            lines.append("")
    return lines, break_types


def _existing_title_errors(paragraph) -> tuple[list[str], list[str]]:
    lines, break_types = _extract_title_lines(paragraph)
    errors = _title_shape_errors(lines)
    invalid_breaks = [break_type for break_type in break_types if break_type != "textWrapping"]
    if invalid_breaks:
        errors.append("标题不得包含分页或分栏换行；请使用普通显式换行。")
    return lines, errors


def _clear_footer(footer) -> None:
    for child in list(footer._element):
        footer._element.remove(child)


def _append_field_run(
    paragraph,
    tag: str,
    attr_name: str | None = None,
    attr_value: str | None = None,
):
    run = paragraph.add_run()
    element = OxmlElement(tag)
    if attr_name is not None and attr_value is not None:
        element.set(qn(attr_name), attr_value)
    run._r.append(element)
    return run, element


def _add_page_field(paragraph) -> None:
    _, begin = _append_field_run(paragraph, "w:fldChar", "w:fldCharType", "begin")
    begin.set(qn("w:dirty"), "true")

    instruction_run = paragraph.add_run()
    instruction = OxmlElement("w:instrText")
    instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instruction.text = r" PAGE \* Arabic "
    instruction_run._r.append(instruction)

    _append_field_run(paragraph, "w:fldChar", "w:fldCharType", "separate")
    paragraph.add_run("1")
    _append_field_run(paragraph, "w:fldChar", "w:fldCharType", "end")

    for run_element in paragraph._p.iter(qn("w:r")):
        _set_run(run_element, PAGE_FONT, PAGE_SIZE_PT, clean_emphasis=True)


def _configure_settings(document) -> None:
    settings = document.settings.element
    _remove_all(settings, "w:mirrorMargins")
    _remove_all(settings, "w:evenAndOddHeaders")
    update_fields = _get_or_add(settings, "w:updateFields")
    update_fields.set(qn("w:val"), "true")


def _configure_sections(document) -> None:
    _configure_settings(document)
    for section in document.sections:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Cm(PAGE_WIDTH_CM)
        section.page_height = Cm(PAGE_HEIGHT_CM)
        section.top_margin = Cm(MARGINS_CM["top"])
        section.bottom_margin = Cm(MARGINS_CM["bottom"])
        section.left_margin = Cm(MARGINS_CM["left"])
        section.right_margin = Cm(MARGINS_CM["right"])
        section.gutter = Cm(0)
        section.footer_distance = Cm(FOOTER_DISTANCE_CM)
        section.different_first_page_header_footer = False
        section.footer.is_linked_to_previous = False

        page_number_type = section._sectPr.find(qn("w:pgNumType"))
        if page_number_type is not None:
            section._sectPr.remove(page_number_type)

        footer = section.footer
        _clear_footer(footer)
        paragraph = footer.add_paragraph()
        _set_paragraph_format(paragraph._p, alignment="center")
        left = paragraph.add_run("- ")
        _set_run(left._r, PAGE_FONT, PAGE_SIZE_PT, clean_emphasis=True)
        _add_page_field(paragraph)
        right = paragraph.add_run(" -")
        _set_run(right._r, PAGE_FONT, PAGE_SIZE_PT, clean_emphasis=True)


def apply_format(
    input_path: Path,
    output_path: Path,
    title_index: int | None,
    title_lines: list[str] | None,
    force: bool = False,
) -> AuditReport:
    if input_path.resolve() == output_path.resolve():
        raise ValueError("输入和输出路径必须不同；本技能不得直接覆盖原稿。")
    if not input_path.is_file():
        raise FileNotFoundError(f"找不到输入文件：{input_path}")
    if input_path.suffix.lower() != ".docx" or output_path.suffix.lower() != ".docx":
        raise ValueError("输入和输出都必须使用 .docx 扩展名。")
    if output_path.exists() and not force:
        raise FileExistsError(f"输出文件已存在：{output_path}；如需替换请显式使用 --force。")

    document = _load_document(input_path)
    _, title = _find_title(document, title_index, allow_empty=bool(title_lines))
    if title_lines:
        shape_errors = _title_shape_errors(title_lines)
        if shape_errors:
            raise ValueError("；".join(shape_errors))
        _replace_title(title, [line.strip() for line in title_lines])
    else:
        existing_lines, shape_errors = _existing_title_errors(title)
        if shape_errors:
            raise ValueError(
                "现有标题断行不合规，请使用多个 --title-line 明确指定：" + "；".join(shape_errors)
            )

    heading_errors = _heading_content_errors(document, title._p)
    if heading_errors:
        raise ValueError("正文标题层级不合规：" + "；".join(heading_errors))

    _configure_styles(document)
    _format_body(document)
    _format_top_level_body_paragraphs(document, title._p)
    _format_title(title)
    _configure_sections(document)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))
    return audit_document(output_path, title_index=title_index, check_fonts=True)


def _value(element, attr: str = "w:val") -> str | None:
    return None if element is None else element.get(qn(attr))


def _check_run(run_element, font_name: str, size_pt: int) -> bool:
    r_pr = run_element.find(qn("w:rPr"))
    if r_pr is None:
        return False
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        return False
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        if r_fonts.get(qn(attr)) != font_name:
            return False
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if r_fonts.get(qn(attr)) is not None:
            return False
    half_points = str(size_pt * 2)
    return (
        _value(r_pr.find(qn("w:sz"))) == half_points
        and _value(r_pr.find(qn("w:szCs"))) == half_points
    )


def _has_text(run_element) -> bool:
    return any((node.text or "") for node in run_element.iter(qn("w:t"))) or any(
        (node.text or "") for node in run_element.iter(qn("w:instrText"))
    )


def _check_spacing(paragraph_element) -> bool:
    p_pr = paragraph_element.find(qn("w:pPr"))
    if p_pr is None:
        return False
    spacing = p_pr.find(qn("w:spacing"))
    return (
        spacing is not None
        and spacing.get(qn("w:before")) == "0"
        and spacing.get(qn("w:after")) == "0"
        and spacing.get(qn("w:line")) == str(LINE_SPACING_PT * 20)
        and spacing.get(qn("w:lineRule")) == "exact"
    )


def _style_element_map(document) -> dict[str, object]:
    return {
        style.get(qn("w:styleId")): style
        for style in document.styles.element.findall(qn("w:style"))
        if style.get(qn("w:styleId"))
    }


def _paragraph_style_chain(document, paragraph_element, style_map=None):
    if style_map is None:
        style_map = _style_element_map(document)
    p_pr = paragraph_element.find(qn("w:pPr"))
    style_id = document.styles["Normal"].style_id
    if p_pr is not None:
        p_style = p_pr.find(qn("w:pStyle"))
        if p_style is not None and p_style.get(qn("w:val")):
            style_id = p_style.get(qn("w:val"))

    visited: set[str] = set()
    while style_id and style_id not in visited:
        visited.add(style_id)
        style = style_map.get(style_id)
        if style is None:
            break
        yield style
        based_on = style.find(qn("w:basedOn"))
        style_id = None if based_on is None else based_on.get(qn("w:val"))


def _effective_paragraph_attributes(
    document,
    paragraph_element,
    tag: str,
    attributes: tuple[str, ...],
    style_map=None,
) -> dict[str, str | None]:
    p_pr = paragraph_element.find(qn("w:pPr"))
    direct = None if p_pr is None else p_pr.find(qn(tag))
    values = {
        attr: None if direct is None else direct.get(qn(attr))
        for attr in attributes
    }
    unresolved = {attr for attr, value in values.items() if value is None}
    for style in _paragraph_style_chain(document, paragraph_element, style_map):
        style_properties = style.find(qn("w:pPr"))
        styled = None if style_properties is None else style_properties.find(qn(tag))
        if styled is None:
            continue
        for attr in tuple(unresolved):
            value = styled.get(qn(attr))
            if value is not None:
                values[attr] = value
                unresolved.remove(attr)
        if not unresolved:
            break
    return values


def _check_body_first_line_indent(paragraph_element, document=None, style_map=None) -> bool:
    if document is None:
        p_pr = paragraph_element.find(qn("w:pPr"))
        indentation = None if p_pr is None else p_pr.find(qn("w:ind"))
        values = {
            attr: None if indentation is None else indentation.get(qn(attr))
            for attr in _INDENT_ATTRIBUTES
        }
    else:
        values = _effective_paragraph_attributes(
            document, paragraph_element, "w:ind", _INDENT_ATTRIBUTES, style_map
        )
    return values["w:firstLineChars"] == str(BODY_FIRST_LINE_CHARS) and all(
        values[attr] in {None, "0"}
        for attr in ("w:firstLine", "w:hanging", "w:hangingChars")
    )


def _find_misordered_section_properties(document) -> list[int]:
    failures: list[int] = []
    for index, p_pr in enumerate(document.element.body.iter(qn("w:pPr"))):
        sect_pr = p_pr.find(qn("w:sectPr"))
        if sect_pr is not None and p_pr.index(sect_pr) != len(p_pr) - 1:
            failures.append(index)
    return failures


def _find_misordered_paragraph_properties(document) -> list[int]:
    failures: list[int] = []
    for index, p_pr in enumerate(document.element.body.iter(qn("w:pPr"))):
        known_orders = [
            _PPR_ORDER[child.tag]
            for child in p_pr
            if child.tag in _PPR_ORDER
        ]
        if any(left > right for left, right in zip(known_orders, known_orders[1:])):
            failures.append(index)
    return failures


def _has_nonzero_first_line_or_hanging_indent(
    paragraph_element, document=None, style_map=None
) -> bool:
    if document is None:
        p_pr = paragraph_element.find(qn("w:pPr"))
        indentation = None if p_pr is None else p_pr.find(qn("w:ind"))
        values = {
            attr: None if indentation is None else indentation.get(qn(attr))
            for attr in _INDENT_ATTRIBUTES
        }
    else:
        values = _effective_paragraph_attributes(
            document, paragraph_element, "w:ind", _INDENT_ATTRIBUTES, style_map
        )
    return any(values[attr] not in {None, "0"} for attr in _INDENT_ATTRIBUTES)


def _has_effective_numbering(document, paragraph_element, style_map=None) -> bool:
    p_pr = paragraph_element.find(qn("w:pPr"))
    if p_pr is not None and p_pr.find(qn("w:numPr")) is not None:
        return True
    for style in _paragraph_style_chain(document, paragraph_element, style_map):
        style_properties = style.find(qn("w:pPr"))
        if style_properties is not None and style_properties.find(qn("w:numPr")) is not None:
            return True
    return False


def _check_alignment(paragraph_element, expected: str) -> bool:
    p_pr = paragraph_element.find(qn("w:pPr"))
    jc = None if p_pr is None else p_pr.find(qn("w:jc"))
    return _value(jc) == expected


def _check_on_off(paragraph_element, tag: str) -> bool:
    p_pr = paragraph_element.find(qn("w:pPr"))
    element = None if p_pr is None else p_pr.find(qn(tag))
    return element is not None and _value(element) in {None, "true", "1"}


def _summarize_indices(indices: list[int]) -> str:
    shown = ", ".join(str(index) for index in indices[:8])
    if len(indices) > 8:
        shown += f" 等 {len(indices)} 处"
    return shown


def _check_installed_fonts(report: AuditReport) -> None:
    if sys.platform == "win32":
        font_dir = Path("C:/Windows/Fonts")
        names = [path.stem.casefold() for path in font_dir.iterdir()] if font_dir.is_dir() else []
        for font_name, hints in FONT_FILE_HINTS.items():
            if not any(any(hint.casefold() in name for hint in hints) for name in names):
                report.warnings.append(f"未在 Windows 字体目录确认字体：{font_name}")
        return

    matcher = shutil.which("fc-match")
    if matcher is None:
        report.warnings.append("当前系统无法自动检查字体安装状态；请在渲染中确认字体未替换。")
        return
    for font_name in FONT_FILE_HINTS:
        result = subprocess.run(
            [matcher, "-f", "%{family}", font_name],
            check=False,
            capture_output=True,
            text=True,
        )
        family = result.stdout.casefold()
        if font_name.casefold() not in family:
            report.warnings.append(f"字体可能被 fontconfig 替换：{font_name} -> {result.stdout.strip()}")


def _twips_match(element, attr: str, expected: int) -> bool:
    if element is None:
        return False
    raw = element.get(qn(attr))
    try:
        return abs(int(raw) - expected) <= 1
    except (TypeError, ValueError):
        return False


def audit_document(path: Path, title_index: int | None, check_fonts: bool) -> AuditReport:
    report = AuditReport(path=str(path.resolve()))
    if not path.is_file():
        report.errors.append(f"找不到文件：{path}")
        return report

    try:
        document = _load_document(path)
    except ValueError as exc:
        report.errors.append(str(exc))
        return report
    try:
        _, title = _find_title(document, title_index)
    except ValueError as exc:
        report.errors.append(str(exc))
        return report

    settings = document.settings.element
    if settings.find(qn("w:mirrorMargins")) is not None:
        report.errors.append("启用了镜像页边距。")
    if settings.find(qn("w:evenAndOddHeaders")) is not None:
        report.errors.append("启用了奇偶页眉页脚。")
    update_fields = settings.find(qn("w:updateFields"))
    if _value(update_fields) not in {"true", "1"}:
        report.errors.append("未设置打开文档时更新字段。")

    for index, section in enumerate(document.sections, start=1):
        page_size = section._sectPr.find(qn("w:pgSz"))
        page_margins = section._sectPr.find(qn("w:pgMar"))
        checks = {
            "page_width": _twips_match(page_size, "w:w", PAGE_SIZE_TWIPS[0]),
            "page_height": _twips_match(page_size, "w:h", PAGE_SIZE_TWIPS[1]),
            "top_margin": _twips_match(page_margins, "w:top", MARGIN_TWIPS["top"]),
            "bottom_margin": _twips_match(page_margins, "w:bottom", MARGIN_TWIPS["bottom"]),
            "left_margin": _twips_match(page_margins, "w:left", MARGIN_TWIPS["left"]),
            "right_margin": _twips_match(page_margins, "w:right", MARGIN_TWIPS["right"]),
            "gutter": _twips_match(page_margins, "w:gutter", MARGIN_TWIPS["gutter"]),
            "footer_distance": _twips_match(
                page_margins, "w:footer", MARGIN_TWIPS["footer"]
            ),
            "portrait": section.orientation in {None, WD_ORIENT.PORTRAIT},
            "continuous_numbering": section._sectPr.find(qn("w:pgNumType")) is None,
            "same_first_page_footer": not section.different_first_page_header_footer,
            "unlinked_footer": not section.footer.is_linked_to_previous,
        }
        failed = [name for name, ok in checks.items() if not ok]
        if failed:
            report.errors.append(f"第 {index} 节版面或页脚设置不合规：{', '.join(failed)}")

    paragraph_elements = _iter_body_paragraph_elements(document)
    section_order_failures = _find_misordered_section_properties(document)
    if section_order_failures:
        report.errors.append(
            "以下段落的分节属性不是 w:pPr 最后一个子元素："
            + _summarize_indices(section_order_failures)
        )
    paragraph_order_failures = _find_misordered_paragraph_properties(document)
    if paragraph_order_failures:
        report.errors.append(
            "以下段落的 w:pPr 子元素顺序不合规："
            + _summarize_indices(paragraph_order_failures)
        )
    spacing_failures = [
        index for index, element in enumerate(paragraph_elements) if not _check_spacing(element)
    ]
    if spacing_failures:
        report.errors.append("以下主体段落不是固定 29 磅且段前后 0：" + _summarize_indices(spacing_failures))

    title_element = title._p
    top_level_paragraph_elements = _iter_top_level_body_paragraph_elements(document)
    style_map = _style_element_map(document)
    heading_by_element = {
        element: heading
        for element in top_level_paragraph_elements
        if element is not title_element
        if (heading := _classify_heading(element)) is not None
    }
    report.errors.extend(_heading_content_errors(document, title_element))
    first_line_failures = [
        index
        for index, element in enumerate(top_level_paragraph_elements)
        if element is not title_element
        and element not in heading_by_element
        and not _check_body_first_line_indent(element, document, style_map)
    ]
    if first_line_failures:
        report.errors.append(
            "以下正文顶层段落未设置首行缩进 2 字符："
            + _summarize_indices(first_line_failures)
        )
    heading_indent_failures = [
        index
        for index, element in enumerate(top_level_paragraph_elements)
        if element in heading_by_element
        and _has_nonzero_first_line_or_hanging_indent(element, document, style_map)
    ]
    if heading_indent_failures:
        report.errors.append(
            "以下正文层级标题不得设置首行缩进或悬挂缩进："
            + _summarize_indices(heading_indent_failures)
        )
    heading_numbering_failures = [
        index
        for index, element in enumerate(top_level_paragraph_elements)
        if element in heading_by_element
        and _has_effective_numbering(document, element, style_map)
    ]
    if heading_numbering_failures:
        report.errors.append(
            "以下正文层级标题带有自动编号属性，可能与字面编号叠加："
            + _summarize_indices(heading_numbering_failures)
        )
    heading_keep_failures = [
        index
        for index, element in enumerate(top_level_paragraph_elements)
        if element in heading_by_element
        and (
            not _check_on_off(element, "w:keepNext")
            or not _check_on_off(element, "w:keepLines")
        )
    ]
    if heading_keep_failures:
        report.errors.append(
            "以下正文层级标题未设置与下段同页或段中不分页："
            + _summarize_indices(heading_keep_failures)
        )
    if _has_nonzero_first_line_or_hanging_indent(title_element, document, style_map):
        report.errors.append("标题不得设置首行缩进或悬挂缩进。")
    if not _check_alignment(title_element, "center"):
        report.errors.append("标题未居中。")
    if not _check_on_off(title_element, "w:keepNext") or not _check_on_off(
        title_element, "w:keepLines"
    ):
        report.errors.append("标题未设置与下段同页或段中不分页。")
    title_lines, title_errors = _existing_title_errors(title)
    report.errors.extend(title_errors)

    body_font_failures: list[int] = []
    heading_font_failures: dict[int, list[int]] = {level: [] for level in range(1, 5)}
    title_font_failures: list[int] = []
    for paragraph_index, paragraph_element in enumerate(paragraph_elements):
        is_title = paragraph_element is title_element
        heading = heading_by_element.get(paragraph_element)
        expected_font = (
            TITLE_FONT
            if is_title
            else HEADING_FONT_BY_LEVEL[heading.level]
            if heading is not None
            else BODY_FONT
        )
        expected_size = TITLE_SIZE_PT if is_title else BODY_SIZE_PT
        for run_element in _iter_paragraph_run_elements(paragraph_element):
            if _has_text(run_element) and not _check_run(run_element, expected_font, expected_size):
                target = (
                    title_font_failures
                    if is_title
                    else heading_font_failures[heading.level]
                    if heading is not None
                    else body_font_failures
                )
                target.append(paragraph_index)
                break
    if title_font_failures:
        report.errors.append("标题字体或字号不合规。")
    if body_font_failures:
        report.errors.append("以下正文段落字体或字号不合规：" + _summarize_indices(body_font_failures))
    for level, failures in heading_font_failures.items():
        if failures:
            report.errors.append(
                f"以下{level}级标题应为 {HEADING_FONT_BY_LEVEL[level]} 三号："
                + _summarize_indices(failures)
            )

    for section_index, section in enumerate(document.sections, start=1):
        footer = section.footer
        instructions = [
            " ".join((node.text or "").split()).upper()
            for node in footer._element.iter(qn("w:instrText"))
        ]
        page_fields = [instruction for instruction in instructions if instruction.startswith("PAGE")]
        visible = "".join(node.text or "" for node in footer._element.iter(qn("w:t")))
        if len(page_fields) != 1:
            report.errors.append(f"第 {section_index} 节页脚 PAGE 字段数量为 {len(page_fields)}，应为 1。")
        elif page_fields[0] != r"PAGE \* ARABIC":
            report.errors.append(f"第 {section_index} 节 PAGE 字段未指定 Arabic 数字格式。")
        field_chars = list(footer._element.iter(qn("w:fldChar")))
        field_types = [node.get(qn("w:fldCharType")) for node in field_chars]
        if field_types != ["begin", "separate", "end"]:
            report.errors.append(f"第 {section_index} 节 PAGE 复杂字段顺序必须为 begin、separate、end。")
        begin_fields = [
            node for node in field_chars if node.get(qn("w:fldCharType")) == "begin"
        ]
        if not begin_fields or begin_fields[0].get(qn("w:dirty")) not in {"true", "1"}:
            report.errors.append(f"第 {section_index} 节 PAGE 字段未标记为需要更新。")

        footer_paragraphs = list(footer._element.iter(qn("w:p")))
        if len(footer_paragraphs) == 1 and field_types == ["begin", "separate", "end"]:
            events: list[tuple[str, str]] = []
            for node in footer_paragraphs[0].iter():
                if node.tag == qn("w:fldChar"):
                    events.append(("field", node.get(qn("w:fldCharType")) or ""))
                elif node.tag == qn("w:instrText"):
                    events.append(("instruction", " ".join((node.text or "").split()).upper()))
                elif node.tag == qn("w:t"):
                    events.append(("text", node.text or ""))

            begin_index = events.index(("field", "begin"))
            separate_index = events.index(("field", "separate"))
            end_index = events.index(("field", "end"))
            field_instruction = [
                value
                for kind, value in events[begin_index + 1 : separate_index]
                if kind == "instruction"
            ]
            cached_result = "".join(
                value
                for kind, value in events[separate_index + 1 : end_index]
                if kind == "text"
            )
            if field_instruction != [r"PAGE \* ARABIC"] or not re.fullmatch(r"\d+", cached_result):
                report.errors.append(f"第 {section_index} 节 PAGE 字段指令或缓存结果位置不正确。")

        if not re.fullmatch(r"- \d+ -", visible):
            report.errors.append(
                f"第 {section_index} 节页脚缓存显示为 {visible!r}，应匹配 '- 1 -' 形式。"
            )
        if len(footer_paragraphs) != 1 or not _check_alignment(footer_paragraphs[0], "center"):
            report.errors.append(f"第 {section_index} 节页码段落不是唯一居中段落。")
        elif not _check_spacing(footer_paragraphs[0]):
            report.errors.append(f"第 {section_index} 节页码段落不是固定 29 磅行距。")
        for run_element in footer._element.iter(qn("w:r")):
            if _has_text(run_element) and not _check_run(run_element, PAGE_FONT, PAGE_SIZE_PT):
                report.errors.append(f"第 {section_index} 节页码字体或字号不合规。")
                break

    if check_fonts:
        _check_installed_fonts(report)
    return report


def _print_report(report: AuditReport, as_json: bool) -> None:
    if as_json:
        print(json.dumps(report.as_dict(), ensure_ascii=False, indent=2))
        return
    print(("PASS" if report.ok else "FAIL") + f": {report.path}")
    for error in report.errors:
        print(f"  ERROR: {error}")
    for warning in report.warnings:
        print(f"  WARNING: {warning}")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="应用并审计用户指定的中文公文 DOCX 格式。"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    apply_parser = subparsers.add_parser("apply", help="格式化 DOCX 并立即审计输出。")
    apply_parser.add_argument("input", type=Path)
    apply_parser.add_argument("output", type=Path)
    apply_parser.add_argument("--title-index", type=int, default=None)
    apply_parser.add_argument(
        "--title-line",
        action="append",
        dest="title_lines",
        help="按语义指定一行标题；多行时重复使用。",
    )
    apply_parser.add_argument(
        "--force",
        action="store_true",
        help="允许替换已经存在的输出文件。",
    )
    apply_parser.add_argument("--json", action="store_true", help="以 JSON 输出审计结果。")

    audit_parser = subparsers.add_parser("audit", help="只读审计现有 DOCX。")
    audit_parser.add_argument("input", type=Path)
    audit_parser.add_argument("--title-index", type=int, default=None)
    audit_parser.add_argument("--json", action="store_true", help="以 JSON 输出结果。")
    audit_parser.add_argument(
        "--skip-font-check",
        action="store_true",
        help="跳过当前运行环境的字体安装检查。",
    )
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        if args.command == "apply":
            report = apply_format(
                args.input,
                args.output,
                title_index=args.title_index,
                title_lines=args.title_lines,
                force=args.force,
            )
            _print_report(report, args.json)
            return 0 if report.ok else 1

        report = audit_document(
            args.input,
            title_index=args.title_index,
            check_fonts=not args.skip_font_check,
        )
        _print_report(report, args.json)
        return 0 if report.ok else 1
    except (OSError, ValueError) as exc:
        if getattr(args, "json", False):
            print(json.dumps({"ok": False, "errors": [str(exc)]}, ensure_ascii=False, indent=2))
        else:
            print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
