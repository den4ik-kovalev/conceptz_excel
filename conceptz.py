"""
Excel VBA - Double Click
https://www.youtube.com/watch?v=vFd2vZw4yJ0

Execute Shell Commands from Excel Cell
https://superuser.com/questions/1220696/execute-shell-commands-from-excel-cell

Python : Reading Large Excel Worksheets using Openpyxl
https://stackoverflow.com/questions/31189727/python-reading-large-excel-worksheets-using-openpyxl
"""

import argparse
import os
import sys
import webbrowser
from pathlib import Path
from typing import Iterator

import openpyxl
from jinja2 import Template
from loguru import logger


ROOT_DIR = Path(os.path.dirname(sys.argv[0]))
LOGGER_PATH = ROOT_DIR / "conceptz.log"
TEMPLATE_PATH = ROOT_DIR / "template.html"
HTML_DIR = ROOT_DIR / "conceptz_html"

CONCEPTS_SHEET_IDX = 0
INFO_SHEET_IDX = 1


logger.add(LOGGER_PATH, format="{time} | {level} | {message}", level="INFO")
os.makedirs(HTML_DIR, exist_ok=True)


def get_excel_rows(
    filepath: Path,
    sheet_idx: int = 0,
    filter_key: str = None,
    filter_value: str = None
) -> Iterator[dict]:

    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.worksheets[sheet_idx]
    if ws.max_row == 0:
        return

    keys = [
        ws.cell(1, column).value
        for column in range(1, ws.max_column + 1)
    ]

    for row in range(2, ws.max_row + 1):
        dct = {
            key: ws.cell(row, column).value
            for column, key in enumerate(keys, start=1)
            if key is not None
        }
        if filter_key:
            if dct[filter_key] == filter_value:
                yield dct
            else:
                continue
        else:
            yield dct


@logger.catch
def main():

    parser = argparse.ArgumentParser(description='Conceptz for Excel')
    parser.add_argument("xls_path", type=str, help="Data file")
    parser.add_argument("concept_name", type=str, help="Concept name")
    args = parser.parse_args()
    logger.info(args.xls_path)
    logger.info(args.concept_name)

    xls_path = Path(args.xls_path)
    concept_name = args.concept_name

    concept_row = get_excel_rows(xls_path, CONCEPTS_SHEET_IDX, filter_key="Name", filter_value=concept_name)
    concept_row = list(concept_row)[0]
    concept_id = concept_row["ID"]
    info_rows = get_excel_rows(xls_path, INFO_SHEET_IDX, filter_key="Concept ID", filter_value=concept_id)
    info_rows = list(info_rows)

    for row in info_rows:
        screenshot = row.get("Screenshot")
        if not screenshot:
            continue
        if "mail.ru" not in screenshot:
            continue
        try:
            xxxx, yyyyyyyyy = screenshot.split('/')[-2:]
        except Exception:
            logger.info(f"Can't parse screenshot link: {screenshot}")
        else:
            row["Screenshot"] = f"https://thumb.cloud.mail.ru/weblink/thumb/xw1/{xxxx}/{yyyyyyyyy}"

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as file:
        template_html = file.read()

    template = Template(template_html)
    rendered_html = template.render(concept_row=concept_row, info_rows=info_rows)
    html_path = HTML_DIR / f"{concept_name}.html"

    with open(html_path, "w", encoding="utf-8") as file:
        file.write(rendered_html)

    webbrowser.open(str(html_path), new=2)


if __name__ == '__main__':
    main()
