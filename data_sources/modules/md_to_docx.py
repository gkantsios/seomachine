#!/usr/bin/env python3
"""
Markdown to Word Document (.docx) Converter

Converts a markdown article file to a formatted .docx file using python-docx.
Parses markdown via the `markdown` library → HTML, then maps HTML tags to
Word paragraph styles.

Usage:
    python3 data_sources/modules/md_to_docx.py drafts/article.md
    python3 data_sources/modules/md_to_docx.py drafts/article.md --output drafts/article.docx
"""

import sys
import re
import argparse
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("Error: python-docx not installed. Run: pip install python-docx")
    sys.exit(1)

try:
    import markdown
    from bs4 import BeautifulSoup, NavigableString, Tag
except ImportError:
    print("Error: Required libraries missing. Run: pip install markdown beautifulsoup4")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _strip_frontmatter(text: str) -> tuple[str, str]:
    """Strip YAML-style frontmatter block (--- ... ---) and return (frontmatter, body)."""
    if text.startswith("---"):
        end = text.find("\n---", 3)
        if end != -1:
            frontmatter = text[3:end].strip()
            body = text[end + 4:].lstrip("\n")
            return frontmatter, body
    return "", text


def _apply_inline_formatting(paragraph, html_element):
    """
    Walk an HTML element's children and add runs with appropriate formatting
    to a python-docx paragraph.
    """
    for child in html_element.children:
        if isinstance(child, NavigableString):
            text = str(child)
            if text:
                paragraph.add_run(text)
        elif isinstance(child, Tag):
            tag = child.name
            text = child.get_text()
            if not text:
                continue

            if tag in ("strong", "b"):
                run = paragraph.add_run(text)
                run.bold = True
            elif tag in ("em", "i"):
                run = paragraph.add_run(text)
                run.italic = True
            elif tag == "code":
                run = paragraph.add_run(text)
                run.font.name = "Courier New"
                run.font.size = Pt(9)
            elif tag == "a":
                # Add link text without hyperlinking (docx hyperlinks require XML manipulation)
                run = paragraph.add_run(text)
                run.font.color.rgb = RGBColor(0x00, 0x56, 0xB3)
            else:
                # Recurse for nested tags (e.g. <strong><em>)
                _apply_inline_formatting(paragraph, child)


def _add_horizontal_rule(doc: Document):
    """Add a thin horizontal rule paragraph."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "AAAAAA")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _process_table(doc: Document, table_tag: Tag):
    """Convert an HTML <table> element to a Word table."""
    rows = table_tag.find_all("tr")
    if not rows:
        return

    # Count columns from first row
    first_row_cells = rows[0].find_all(["th", "td"])
    num_cols = len(first_row_cells)
    if num_cols == 0:
        return

    word_table = doc.add_table(rows=0, cols=num_cols)
    word_table.style = "Table Grid"

    for i, row in enumerate(rows):
        cells = row.find_all(["th", "td"])
        word_row = word_table.add_row()
        for j, cell in enumerate(cells):
            if j >= num_cols:
                break
            p = word_row.cells[j].paragraphs[0]
            _apply_inline_formatting(p, cell)
            if cell.name == "th" or i == 0:
                for run in p.runs:
                    run.bold = True

    doc.add_paragraph()  # spacing after table


def _process_list(doc: Document, list_tag: Tag, level: int = 0):
    """Recursively process <ul> or <ol> tags into Word list paragraphs."""
    is_ordered = list_tag.name == "ol"
    for item in list_tag.find_all("li", recursive=False):
        # Check for nested lists inside this item
        nested = item.find(["ul", "ol"])
        # Get text without nested list text
        if nested:
            item_text = "".join(
                str(c) for c in item.children
                if not (isinstance(c, Tag) and c.name in ("ul", "ol"))
            )
            item_soup = BeautifulSoup(item_text, "html.parser")
        else:
            item_soup = item

        style = "List Number" if is_ordered else "List Bullet"
        # Indent nested lists
        if level > 0:
            style = "List Bullet 2" if not is_ordered else "List Number 2"

        try:
            p = doc.add_paragraph(style=style)
        except KeyError:
            p = doc.add_paragraph()

        _apply_inline_formatting(p, item_soup)

        # Process nested list
        if nested:
            _process_list(doc, nested, level + 1)


def _process_blockquote(doc: Document, bq_tag: Tag):
    """Render a <blockquote> with indented, italic style."""
    for child in bq_tag.children:
        if isinstance(child, NavigableString):
            text = str(child).strip()
            if text:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.5)
                run = p.add_run(text)
                run.italic = True
        elif isinstance(child, Tag):
            if child.name == "p":
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.5)
                _apply_inline_formatting(p, child)
                for run in p.runs:
                    run.italic = True
            elif child.name in ("ul", "ol"):
                _process_list(doc, child)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.5)
                _apply_inline_formatting(p, child)
                for run in p.runs:
                    run.italic = True


# ---------------------------------------------------------------------------
# Main converter
# ---------------------------------------------------------------------------

def convert_md_to_docx(md_path: str, output_path: str | None = None) -> str:
    """
    Convert a markdown file to a .docx file.

    Args:
        md_path: Path to the input .md file.
        output_path: Optional output path. Defaults to same name with .docx extension.

    Returns:
        Path to the created .docx file.
    """
    md_file = Path(md_path)
    if not md_file.exists():
        raise FileNotFoundError(f"Markdown file not found: {md_path}")

    if output_path is None:
        output_path = str(md_file.with_suffix(".docx"))

    raw_text = md_file.read_text(encoding="utf-8")

    # Strip frontmatter
    frontmatter, body = _strip_frontmatter(raw_text)

    # Convert markdown to HTML
    md_extensions = ["tables", "fenced_code", "nl2br", "sane_lists"]
    html = markdown.markdown(body, extensions=md_extensions)

    # Parse HTML
    soup = BeautifulSoup(html, "html.parser")

    # Build Word document
    doc = Document()

    # --- Document styles ---
    # Adjust Normal style for readability
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)

    # --- Frontmatter as document property block (italicised, small) ---
    if frontmatter:
        for line in frontmatter.splitlines():
            line = line.strip()
            if line:
                p = doc.add_paragraph()
                run = p.add_run(line)
                run.font.size = Pt(9)
                run.italic = True
                run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        _add_horizontal_rule(doc)

    # --- Process HTML elements ---
    for element in soup.children:
        if isinstance(element, NavigableString):
            text = str(element).strip()
            if text:
                doc.add_paragraph(text)
            continue

        if not isinstance(element, Tag):
            continue

        tag = element.name

        # Headings
        if tag == "h1":
            p = doc.add_heading(level=1)
            p.clear()
            _apply_inline_formatting(p, element)

        elif tag == "h2":
            p = doc.add_heading(level=2)
            p.clear()
            _apply_inline_formatting(p, element)

        elif tag == "h3":
            p = doc.add_heading(level=3)
            p.clear()
            _apply_inline_formatting(p, element)

        elif tag == "h4":
            p = doc.add_heading(level=4)
            p.clear()
            _apply_inline_formatting(p, element)

        # Paragraphs
        elif tag == "p":
            # Check if it's purely a <hr> stand-in
            inner = element.get_text(strip=True)
            if inner in ("---", "***", "___"):
                _add_horizontal_rule(doc)
            else:
                p = doc.add_paragraph()
                _apply_inline_formatting(p, element)

        # Lists
        elif tag in ("ul", "ol"):
            _process_list(doc, element)

        # Blockquotes (Key Takeaways, callouts)
        elif tag == "blockquote":
            _process_blockquote(doc, element)

        # Tables
        elif tag == "table":
            _process_table(doc, element)

        # Code blocks
        elif tag == "pre":
            code = element.find("code")
            text = code.get_text() if code else element.get_text()
            p = doc.add_paragraph()
            run = p.add_run(text)
            run.font.name = "Courier New"
            run.font.size = Pt(9)
            p.paragraph_format.left_indent = Inches(0.25)

        # Horizontal rules
        elif tag == "hr":
            _add_horizontal_rule(doc)

    doc.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Convert a markdown article to a Word .docx file."
    )
    parser.add_argument("input", help="Path to the input .md file")
    parser.add_argument(
        "--output", "-o",
        help="Output .docx path (default: same name as input with .docx extension)",
        default=None,
    )
    args = parser.parse_args()

    try:
        out = convert_md_to_docx(args.input, args.output)
        print(f"Saved: {out}")
    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
