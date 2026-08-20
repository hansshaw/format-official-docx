from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "scripts" / "official_docx.py"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def text_of(node) -> str:
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))


def value(node, attribute: str = "val") -> str | None:
    return node.get(f"{{{NS['w']}}}{attribute}") if node is not None else None


class OfficialDocxTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = Path(tempfile.mkdtemp(prefix="official-docx-test-"))

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def run_cli(self, *args: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(CLI), *args],
            cwd=ROOT,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )

    def build_draft(self, path: Path, invalid_heading: str | None = None) -> None:
        doc = Document()
        doc.add_paragraph("关于开展年度重点工作的通知")
        doc.add_paragraph("各有关单位：")
        doc.add_paragraph("现就有关事项通知如下。")
        doc.add_paragraph("一、总体要求")
        doc.add_paragraph("坚持问题导向，压实工作责任。")
        doc.add_paragraph("（一）重点任务。")
        doc.add_paragraph("1.加强统筹协调。")
        doc.add_paragraph("（1）完善工作机制。")
        if invalid_heading:
            doc.add_paragraph(invalid_heading)
        doc.add_paragraph("请按时反馈办理情况。")
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "（一）表格内文本不应作为正文标题"
        doc.save(path)

    def read_document_xml(self, path: Path):
        from lxml import etree

        with zipfile.ZipFile(path) as package:
            return etree.fromstring(package.read("word/document.xml"))

    def read_footer_xml(self, path: Path):
        from lxml import etree

        with zipfile.ZipFile(path) as package:
            footer_names = sorted(
                name for name in package.namelist() if name.startswith("word/footer") and name.endswith(".xml")
            )
            self.assertTrue(footer_names)
            return [etree.fromstring(package.read(name)) for name in footer_names]

    def test_apply_formats_body_headings_and_page_field(self) -> None:
        draft = self.temp_dir / "draft.docx"
        output = self.temp_dir / "final.docx"
        self.build_draft(draft)
        result = self.run_cli(
            "apply",
            str(draft),
            str(output),
            "--title-line",
            "关于开展年度重点工作的",
            "--title-line",
            "通知",
            "--json",
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        report = json.loads(result.stdout)
        self.assertTrue(report["ok"], report)

        root = self.read_document_xml(output)
        paragraphs = root.xpath("//w:body//w:p", namespaces=NS)
        by_text = {text_of(p): p for p in paragraphs}

        page_size = root.xpath("//w:body/w:sectPr/w:pgSz", namespaces=NS)[0]
        page_margins = root.xpath("//w:body/w:sectPr/w:pgMar", namespaces=NS)[0]
        self.assertEqual(value(page_size, "w"), "11906")
        self.assertEqual(value(page_size, "h"), "16838")
        self.assertEqual(value(page_margins, "top"), "2098")
        self.assertEqual(value(page_margins, "bottom"), "1984")
        self.assertEqual(value(page_margins, "left"), "1587")
        self.assertEqual(value(page_margins, "right"), "1474")

        title = by_text["关于开展年度重点工作的通知"]
        title_pr = title.xpath("./w:pPr", namespaces=NS)[0]
        self.assertEqual(value(title_pr.xpath("./w:jc", namespaces=NS)[0]), "center")
        title_fonts = title.xpath(".//w:rPr/w:rFonts", namespaces=NS)[0]
        for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
            self.assertEqual(value(title_fonts, attribute), "方正小标宋简体")
        self.assertEqual(value(title.xpath(".//w:rPr/w:sz", namespaces=NS)[0]), "44")

        expected_fonts = {
            "一、总体要求": "黑体",
            "（一）重点任务。": "楷体_GB2312",
            "1.加强统筹协调。": "仿宋_GB2312",
            "（1）完善工作机制。": "仿宋_GB2312",
        }
        for heading, font in expected_fonts.items():
            paragraph = by_text[heading]
            r_fonts = paragraph.xpath(".//w:rPr/w:rFonts", namespaces=NS)[0]
            for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
                self.assertEqual(value(r_fonts, attribute), font, heading)
            ind_nodes = paragraph.xpath("./w:pPr/w:ind", namespaces=NS)
            if ind_nodes:
                self.assertIsNone(value(ind_nodes[0], "firstLineChars"), heading)
            self.assertTrue(paragraph.xpath("./w:pPr/w:keepNext", namespaces=NS), heading)

        body = by_text["各有关单位："]
        body_ind = body.xpath("./w:pPr/w:ind", namespaces=NS)[0]
        self.assertEqual(value(body_ind, "firstLineChars"), "200")
        for paragraph in (title, *expected_fonts.keys(), "各有关单位："):
            node = by_text[paragraph] if isinstance(paragraph, str) else paragraph
            spacing = node.xpath("./w:pPr/w:spacing", namespaces=NS)[0]
            self.assertEqual(value(spacing, "before"), "0")
            self.assertEqual(value(spacing, "after"), "0")
            self.assertEqual(value(spacing, "line"), "580")
            self.assertEqual(value(spacing, "lineRule"), "exact")

        footers = self.read_footer_xml(output)
        self.assertTrue(
            any(footer.xpath(".//w:fldChar[@w:fldCharType='begin']", namespaces=NS) for footer in footers)
        )
        self.assertTrue(
            any(footer.xpath(".//w:instrText[contains(., 'PAGE') ]", namespaces=NS) for footer in footers)
        )
        self.assertTrue(any(text_of(footer) == "- 1 -" for footer in footers))
        page_fonts = footers[0].xpath(".//w:rPr/w:rFonts", namespaces=NS)
        self.assertTrue(page_fonts)
        for r_fonts in page_fonts:
            self.assertEqual(value(r_fonts, "eastAsia"), "宋体")

    def test_audit_rejects_invalid_headings_without_creating_output(self) -> None:
        invalid_cases = (
            ("1.缺少句末标点", "句末必须有标点"),
            ("一、一级标题。", "一级标题句末不得有标点"),
            ("(一)错误括号。", "二级标题应使用全角括号"),
        )
        for index, (heading, expected_error) in enumerate(invalid_cases):
            with self.subTest(heading=heading):
                draft = self.temp_dir / f"draft-{index}.docx"
                output = self.temp_dir / f"should-not-exist-{index}.docx"
                self.build_draft(draft, heading)
                result = self.run_cli("apply", str(draft), str(output))
                self.assertNotEqual(result.returncode, 0)
                self.assertFalse(output.exists())
                self.assertIn(expected_error, result.stdout + result.stderr)

    def test_apply_does_not_overwrite_without_force(self) -> None:
        draft = self.temp_dir / "draft.docx"
        output = self.temp_dir / "final.docx"
        self.build_draft(draft)
        output.write_bytes(b"sentinel")
        result = self.run_cli("apply", str(draft), str(output))
        self.assertEqual(result.returncode, 2)
        self.assertEqual(output.read_bytes(), b"sentinel")

    def test_audit_json_can_skip_font_check(self) -> None:
        draft = self.temp_dir / "draft.docx"
        output = self.temp_dir / "final.docx"
        self.build_draft(draft)
        applied = self.run_cli("apply", str(draft), str(output))
        # Missing fonts are warnings; structural validation remains deterministic on CI.
        self.assertEqual(applied.returncode, 0, applied.stdout + applied.stderr)
        audited = self.run_cli("audit", str(output), "--skip-font-check", "--json")
        self.assertEqual(audited.returncode, 0, audited.stdout + audited.stderr)
        self.assertTrue(json.loads(audited.stdout)["ok"])


if __name__ == "__main__":
    unittest.main()
