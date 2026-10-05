"""Export a four-column Jusen product mapping workbook as Markdown.

Usage:
    python export_mapping_markdown.py input.xlsx output.md
"""

from pathlib import Path
import sys

import openpyxl


def cell(value):
    return "" if value is None else str(value).replace("|", "\\|").replace("\n", " ").strip()


def main(source_path: str, output_path: str) -> None:
    workbook = openpyxl.load_workbook(source_path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("工作簿没有可导出的内容")
    headers = [cell(value) for value in rows[0][:4]]
    if headers != ["产品名称", "品类", "细分品类", "系列名"]:
        raise ValueError(f"表头不符合预期：{headers}")
    lines = [
        "# 炬森产品三级分类对照表（2026）",
        "",
        "用于产品名称匹配时的分类参考。名称存在版本差异时，型号、规格和人工备注优先于文本相似度。",
        "",
        "| 产品名称 | 品类 | 细分品类 | 系列名 |",
        "| --- | --- | --- | --- |",
    ]
    for row in rows[1:]:
        if row[0] is not None:
            lines.append("| " + " | ".join(cell(value) for value in row[:4]) + " |")
    Path(output_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("用法：python export_mapping_markdown.py 输入.xlsx 输出.md")
    main(sys.argv[1], sys.argv[2])
