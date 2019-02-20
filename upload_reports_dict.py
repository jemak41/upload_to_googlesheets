import string
from datetime import datetime as dt
import pandas as pd
import sys

import sheets.getcred as gc

value_input_option = 'user_entered'
reportRange = 'A1:Z10000'

def getStartRow(reportRange, spreadsheet_id, sheetsName):
    response = gc.service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sheetsName + reportRange).execute()
    if 'values' not in response:
        return 1
    return len(response['values']) + 1

def getEndRow(df, start):
    return len(df.index) + start + 1

def getRange(df, reportRange, spreadsheet_id, sheetsName):
    col = len(df.columns)
    start = getStartRow(reportRange, spreadsheet_id, sheetsName) + 1
    rang = sheetsName + 'A' + str(start)
    rang += ':'
    rang += list(string.ascii_uppercase)[col - 1]
    end = getEndRow(df, start) + 1
    rang += str(end)
    return rang

def getValues(df, dfTitle):
    title = [dfTitle]
    headers = df.columns.values.tolist()
    values = df.values.tolist()
    values.insert(0, title)
    values.insert(1, headers)
    return values

def printReport(df, dfTitle, spreadsheet_id, sheetsName):
    range_name = getRange(df, reportRange, spreadsheet_id, sheetsName)
    values = getValues(df, dfTitle)
    data = [
        {
            'range': range_name,
            'values': values
        }
    ]
    body = {
        'valueInputOption': value_input_option,
        'data': data
    }
    result = gc.service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def printAnyReport(dict, spreadsheet_id, sheetsName='Sheet1!'):
    for key, value in dict.items():
        if not isinstance(value, pd.DataFrame):
            value = pd.DataFrame([value])
        printReport(value, 'This is the ' + key + ' report, generated today (' + str(dt.now()) + ')', spreadsheet_id, sheetsName)

