"""
Author: Mukul Hase

Set of custom functions for modifying formatting of a google spreadsheet
"""


def add_formatting(sheet_id, col, total_rows, options):
    new_options = list(map(lambda op: {
        "userEnteredValue": op
    }, options))
    requests = [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": total_rows - 1,
                    "startColumnIndex": col - 1,
                    "endColumnIndex": col
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": new_options
                    },
                    "inputMessage": "Value must be one of " + ",".join(options),
                    "strict": True
                }
            }
        }]
    return requests


def set_col_size(sheet_id, col, size):
    return [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": col - 1,
                    "endIndex": col
                },
                "properties": {
                    "pixelSize": size
                },
                "fields": "pixelSize"
            }
        }
    ]
