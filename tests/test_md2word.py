"""Tests for md2word conversion, styles, math, and PDF."""

from __future__ import annotations

import re
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from md2word.converter import (
    find_pandoc,
    markdown_to_docx,
    markdown_to_html,
    markdown_to_pdf,
    normalize_markdown,
)
from md2word.mathprep import (
    prepare_markdown_layout,
    prepare_markdown_math,
    strip_thematic_breaks,
)
from md2word import clipboard
from md2word.i18n import set_lang, t
from md2word.wordcom import word_installed


SAMPLE = r"""## 正态分布

连续随机变量若服从正态分布，常写成 (N(\mu,\sigma^2))：

- 均值：(\mu)
- 标准差：(\sigma)
- 标准化：(\hat{z})

密度函数：

[

f(x)=\frac{1}{\sigma\sqrt{2\pi}}\exp\left(-\frac{(x-\mu)^2}{2\sigma^2}\right)

]

这是密度，不是单点概率。

---

## 一元二次方程

若 $\Delta=b^2-4ac>0$，则

[

x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}

]
"""


def _docx_xml(path: Path) -> tuple[str, str]:
    with zipfile.ZipFile(path) as z:
        document = z.read("word/document.xml").decode("utf-8")
        styles = z.read("word/styles.xml").decode("utf-8")
    return document, styles


class ClipboardApiTests(unittest.TestCase):
    def test_public_helpers_exist(self) -> None:
        self.assertTrue(callable(clipboard.get_text))
        self.assertTrue(callable(clipboard.set_rich_for_word))
        self.assertTrue(clipboard.paste_shortcut())

    def test_i18n_english_and_chinese(self) -> None:
        from md2word.i18n import lang as current

        prev = current()
        try:
            set_lang("en")
            self.assertEqual(t("ready"), "Ready")
            set_lang("zh")
            self.assertEqual(t("ready"), "就绪")
        finally:
            set_lang(prev)


class MathPrepTests(unittest.TestCase):
    def test_wraps_paren_latex(self) -> None:
        out = prepare_markdown_math(r"均值 (\mu) 与 (\hat{x})")
        self.assertIn(r"$\mu$", out)
        self.assertIn("hat", out)
        self.assertIn("$", out)

    def test_wraps_bracket_block(self) -> None:
        src = "前文\n\n[\n\n\\Delta=b^2-4ac\n\n]\n\n后文\n"
        out = prepare_markdown_math(src)
        self.assertIn("$$", out)
        self.assertIn(r"\Delta=b^2-4ac", out)
        self.assertNotRegex(out, r"(?m)^\[$")

    def test_keeps_existing_dollar_math(self) -> None:
        src = r"若 $\Delta>0$ 则有两个实根。"
        out = prepare_markdown_math(src)
        self.assertIn(r"$\Delta>0$", out)

    def test_does_not_touch_code(self) -> None:
        src = "code `\\alpha` and\n```\n\\beta\n```\n"
        out = prepare_markdown_math(src)
        self.assertIn("`\\alpha`", out)
        self.assertIn("```\n\\beta\n```", out)

    def test_strips_thematic_breaks(self) -> None:
        src = "段一\n\n---\n\n段二\n\n***\n\n段三\n"
        out = strip_thematic_breaks(src)
        self.assertNotRegex(out, r"(?m)^---$")
        self.assertNotRegex(out, r"(?m)^\*\*\*$")
        self.assertIn("段一", out)
        self.assertIn("段二", out)

    def test_does_not_double_dollar_parenthetical_math(self) -> None:
        src = r"**表2. 实验 ($M^{tar}(pProduct) = 10$)**"
        out = prepare_markdown_math(src)
        self.assertIn(r"$M^{tar}(pProduct) = 10$", out)
        self.assertNotIn("$$", out)
        self.assertNotIn("($M", out)

    def test_rewrites_tag_for_pandoc(self) -> None:
        src = r"$$\lambda_B = 1 \tag{1}$$"
        out = prepare_markdown_math(src)
        self.assertNotIn(r"\tag", out)
        self.assertIn("(1)", out)

    def test_layout_blank_line_before_table_and_list(self) -> None:
        src = (
            "**表1. 标题**\n"
            "| a | b |\n"
            "| --- | --- |\n"
            "| 1 | 2 |\n\n"
            "问题如下：\n"
            "1. 第一问\n"
            "2. 第二问\n"
        )
        out = prepare_markdown_layout(src)
        self.assertRegex(out, r"标题\*\*\n\n\| a \| b \|")
        self.assertRegex(out, r"问题如下：\n\n1\. 第一问")


class NormalizeTests(unittest.TestCase):
    def test_normalize_sample(self) -> None:
        out = normalize_markdown(SAMPLE)
        self.assertIn("$$", out)
        self.assertIn(r"\mu", out)
        self.assertNotRegex(out, r"(?m)^---+$")
        self.assertIn(r"$\mu$", out)


@unittest.skipUnless(find_pandoc(), "pandoc not installed")
class DocxConvertTests(unittest.TestCase):
    def test_docx_has_omml_and_body_styles(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "out.docx"
            res = markdown_to_docx(SAMPLE, path)
            self.assertTrue(res.success, res.message)
            self.assertTrue(path.is_file())
            document, styles = _docx_xml(path)

            self.assertGreater(document.count("oMath"), 3, "expected Word OMML equations")
            self.assertNotIn("pBdr", document)
            self.assertNotRegex(document, r"(?m)^---")

            # leftover raw latex commands should not appear as plain text
            text = re.sub(r"<m:oMath[\s\S]*?</m:oMath>", " ", document)
            text = re.sub(r"<[^>]+>", "", text)
            self.assertNotIn(r"\sigma", text)
            self.assertNotIn(r"\hat", text)

            self.assertIn('w:firstLineChars="200"', document + styles)
            self.assertIn('w:jc w:val="both"', document + styles)
            self.assertIn('w:line="360"', document + styles)

    def test_html_has_no_hr_and_has_math(self) -> None:
        res = markdown_to_html(SAMPLE)
        self.assertTrue(res.success, res.message)
        html = res.content or ""
        self.assertIn("math", html.lower())
        self.assertIn("text-indent: 2em", html)
        self.assertIn("text-align: justify", html)
        self.assertIn("line-height: 1.5", html)

    def test_tables_lists_fonts_and_equation_tags(self) -> None:
        src = r"""## 标题 Heading

正文先行，Times New Roman 与中文混排。

问题如下：
1. 第一项需要换行后仍按正文首行缩进。
2. 第二项。

**表1. 质量参数 ($M^{tar}=10$)**
| 名称 | 值 $\alpha$ |
| :--- | :---------: |
| A | 0.85 |
| B | 0.10 |

公式：
$$\lambda_B = \frac{1-q_B}{\alpha_B} \tag{1}$$

## 参考文献

1. Zhang C. Robust Control // IEEE. 2024. V.21. P. 1–10.
2. Сочнев А.Н. 网络模型. 2020.
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "layout.docx"
            res = markdown_to_docx(src, path)
            self.assertTrue(res.success, res.message)
            document, styles = _docx_xml(path)
            self.assertIn("<w:tbl", document)
            self.assertNotIn(r"\tag", document)
            self.assertGreater(document.count("oMath"), 1)
            self.assertIn('w:ascii="Times New Roman"', document + styles)
            self.assertIn('w:jc w:val="both"', document + styles)
            numbered = re.findall(
                r"<w:numPr>[\s\S]*?</w:p>",
                document,
            )
            self.assertGreaterEqual(len(numbered), 2)
            for block in numbered:
                self.assertIn('w:firstLineChars="200"', block)
                self.assertIn('w:leftChars="0"', block)
                self.assertIn('w:rightChars="0"', block)
                self.assertNotIn("w:hanging", block)
            self.assertNotRegex(document, r"\$\$\\lambda")

    def test_list_continuations_after_display_math_are_not_renumbered(self) -> None:
        src = """## 附录

1. 解得：
   $$x=1$$
   因此继续推导。
2. 第二点：
   $$y=2$$
   得证。
"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "list.docx"
            res = markdown_to_docx(src, path)
            self.assertTrue(res.success, res.message)
            document, _ = _docx_xml(path)
            numbered = re.findall(r"<w:numPr>[\s\S]*?</w:p>", document)
            self.assertEqual(len(numbered), 2)
            self.assertIn("因此继续推导", document)
            self.assertIn("得证", document)
            self.assertIn('w:firstLineChars="200"', document)


class PdfExportTests(unittest.TestCase):
    def test_pdf_export_if_word_present(self) -> None:
        if not word_installed():
            self.skipTest("Microsoft Word not installed")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "out.pdf"
            res = markdown_to_pdf(SAMPLE, path)
            self.assertTrue(res.success, res.message)
            self.assertTrue(path.is_file())
            self.assertGreater(path.stat().st_size, 1000)


if __name__ == "__main__":
    unittest.main(verbosity=2)
