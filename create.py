import argparse
import csv
import glob
import os
import time
from collections import namedtuple

import gspread
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials
from tqdm import tqdm

from utils import add_formatting, set_col_size

Column = namedtuple("Column", ["col_no", "cellFormat", "col_size", "options"])

columns = [
    Column(
        1,
        cellFormat(
            backgroundColor=color(1, 0.9, 0.9),
            textFormat=textFormat(bold=True, foregroundColor=color(1, 0, 1)),
            horizontalAlignment='CENTER',
            borders={
                "right": {
                    "style": "DOUBLE",
                },
                "bottom": {
                    "style": "SOLID_MEDIUM"
                }
            }
        ),
        50,
        None
    ),
    Column(
        2,
        cellFormat(
            backgroundColor=color(1, 1, 1),
            textFormat=textFormat(foregroundColor=color(0, 0, 0)),
            horizontalAlignment='CENTER',
            wrapStrategy='WRAP',
            borders={
                "bottom": {
                    "style": "SOLID_MEDIUM"
                }
            },
            padding={
                "top": 10
            }
        ),
        700,
        None
    ),
    Column(
        3,
        cellFormat(
            backgroundColor=color(1, 0.9, 0.9),
            textFormat=textFormat(bold=True, foregroundColor=color(1, 0, 1)),
            horizontalAlignment='CENTER',
            borders={
                "bottom": {
                    "style": "SOLID_MEDIUM"
                }
            }
        ),
        None,
        [
            "opening",
            "minor topic start",
            "minor topic end",
            "major topic start",
            "major topic end",
            "off topic start",
            "off topic end",
            "closing",
        ]
    ),
    Column(
        4,
        cellFormat(
            backgroundColor=color(1, 0.9, 0.9),
            textFormat=textFormat(bold=True, foregroundColor=color(1, 0, 1)),
            horizontalAlignment='CENTER',
            borders={
                "bottom": {
                    "style": "SOLID_MEDIUM"
                }
            }
        ),
        None,
        [
            "opening",
            "minor topic start",
            "minor topic end",
            "major topic start",
            "major topic end",
            "off topic start",
            "off topic end",
            "closing",
        ]
    ),
    Column(
        5,
        cellFormat(
            backgroundColor=color(1, 0.9, 0.9),
            textFormat=textFormat(bold=True, foregroundColor=color(0.263, 0.671, 0.788)),
            horizontalAlignment='CENTER',
            borders={
                "bottom": {
                    "style": "SOLID_MEDIUM"
                }
            }
        ),
        None,
        None
    ),
]

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

parser = argparse.ArgumentParser()
parser.add_argument("-json_keyfile", default="My Project 54048-3954a2c146cd.json")
parser.add_argument("-spread_sheet_id", default='1lw5CBhCxTCVaw-bLbQnxJdssyxtZb076f0_nUnOeamI')
parser.add_argument("-data", default="../")
command = parser.parse_args()

# Login to service account
credentials = ServiceAccountCredentials.from_json_keyfile_name(command.json_keyfile, scope)

# Use credentials to create gspread instance
gc = gspread.authorize(credentials)

# Open google spread sheet
sh = gc.open_by_key(command.spread_sheet_id)

data_path = command.data
files = glob.glob(os.path.join(data_path, "*.txt"))[1:5]


def add_table_to_sheet(ws, rows):
    headers_printed = False
    cellList = []
    column_count = 1
    row_count = 1
    for row in rows:
        for data in row.keys():
            if not headers_printed:
                for header in list(row.keys()):
                    cellList.append((column_count, row_count, header))
                    row_count += 1
                row_count = 1
                column_count += 1
                headers_printed = True
            cellList.append((column_count, row_count, row[data]))
            row_count += 1
        column_count += 1
        row_count = 1
    cellListNew = list(map(lambda tup: gspread.models.Cell(*tup), cellList))
    ws.update_cells(cellListNew)

list_of_ws = []
for file_path in tqdm(files):
    fh = open(file_path, 'rt')
    reader = csv.DictReader(fh, delimiter="|",
                            fieldnames=["Person", "Conversation", "Tag",
                                        # Extra Fields added (default to None)
                                        "Custom Tag", "Custom Tag 2", "Topic Description"])
    rows = list(reader)
    for row in rows:
        del row["Tag"]
    total = len(rows)
    fh.seek(0)

    name_of_sheet = os.path.splitext(os.path.basename(file_path))[0]

    try:
        ws = sh.add_worksheet(title=name_of_sheet, rows=total + 1, cols="3")
    except:
        # Skipping sheet cause it's already created
        continue
        pass

    list_of_ws.append(ws)
    add_table_to_sheet(ws, rows)

    time.sleep(2)

    # Freeze header
    set_frozen(ws, rows=1)

    # Freeze first column?
    # set_frozen(ws, cols=1)

    time.sleep(1)

    ranges = []
    for col in columns:
        def column_to_range(col_no):
            return chr(ord('A')+col_no-1)
        ranges.append(
            "%s1:%s%s" % (
                column_to_range(col.col_no),
                column_to_range(col.col_no),
                ws.row_count))

    # Apply cell formats
    format_cell_ranges(ws, list(zip(ranges, [col.cellFormat for col in columns])))

# Batching all custom style application requests in one single huge update
requests = []
for ws in list_of_ws:
    for col in columns:
        if col.col_size:
            requests += set_col_size(ws.id, col.col_no, col.col_size)
        if col.options:
            requests += add_formatting(ws.id, col.col_no, ws.row_count, col.options)

if len(requests)>0:
    sh.batch_update({
        "requests": requests
    })
