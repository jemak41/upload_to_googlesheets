import sys
import pandas as pd
import string


import sheets.getcred as gc

reportRange = 'A1:Z100000'
value_input_option = 'user_entered'

def getStartRow(reportRange, spreadsheet_id, sheetsName):
    response = gc.service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sheetsName + reportRange).execute()
    if 'values' not in response:
        return 1
    return len(response['values']) + 1

def getEndRow(df, start):
    return len(df.index) + start

def getMaxIDfromSheet(spreadsheet_id, sheetsName):
    response = gc.service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sheetsName + reportRange).execute()
    if 'values' not in response:
        return 0
    df = pd.DataFrame(response['values'])
    df[0] = pd.to_numeric(df[0])
    max = df.iloc[:,0].max()
    return max


def getRange(df, reportRange, spreadsheet_id, sheetsName):
    col = len(df.columns)
    start = getStartRow(reportRange, spreadsheet_id, sheetsName)
    rang = sheetsName + 'A' + str(start)
    rang += ':'
    rang += list(string.ascii_uppercase)[col - 1]

    end = getEndRow(df, start)
    rang += str(end)
    return rang


def getValues(df, MaxID):
    df1 = df[df.ID > MaxID]
    values = df1.values.tolist()
    return values

def printReport(df, spreadsheet_id, sheetsName):
    MaxID = getMaxIDfromSheet(spreadsheet_id, sheetsName)
    range_name = getRange(df, reportRange, spreadsheet_id, sheetsName)
    values = getValues(df, MaxID)
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

def printAnyReport(df, spreadsheet_id, sheetName):
    for x in df.columns:
        df[str(x)] = df[str(x)].astype(str)
    printReport(df, spreadsheet_id, sheetName)

