import argparse
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_block_items(parent):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)


def paragraph_to_markdown(paragraph):
    text = paragraph.text.strip()
    if not text:
        return ""

    style_name = (paragraph.style.name or "").lower()
    if style_name.startswith("heading"):
        level = 1
        parts = style_name.split()
        if len(parts) > 1 and parts[1].isdigit():
            level = max(1, min(6, int(parts[1])))
        return f"{'#' * level} {text}"

    return text


def table_to_markdown(table):
    rows = []
    for row in table.rows:
        cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
        rows.append(cells)

    if not rows:
        return ""

    col_count = max(len(row) for row in rows)
    normalized_rows = []
    for row in rows:
        if len(row) < col_count:
            row = row + [""] * (col_count - len(row))
        normalized_rows.append(row)

    header = normalized_rows[0]
    separator = ["---"] * col_count

    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(separator) + " |",
    ]

    for row in normalized_rows[1:]:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)


def convert_docx_to_markdown(input_docx, output_md):
    document = Document(input_docx)
    chunks = []

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            paragraph_md = paragraph_to_markdown(block)
            if paragraph_md:
                chunks.append(paragraph_md)
        elif isinstance(block, Table):
            table_md = table_to_markdown(block)
            if table_md:
                chunks.append(table_md)

    markdown_content = "\n\n".join(chunks).strip() + "\n"
    with open(output_md, "w", encoding="utf-8") as file:
        file.write(markdown_content)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert report DOCX to Markdown for debugging.")
    parser.add_argument("--input", default="Final_Shiksha_Report.docx", help="Input .docx file path")
    parser.add_argument("--output", default="Final_Shiksha_Report.md", help="Output .md file path")
    args = parser.parse_args()

    convert_docx_to_markdown(args.input, args.output)
    print(f"✅ Markdown generated: {args.output}")
