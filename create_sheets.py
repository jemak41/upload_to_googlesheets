import sys
import getcred as gc

def createSheet(spreadsheet_id, sheetName):
    batch_update_spreadsheet_request_body = {
        'requests': [
            {
                'addSheet': {
                    'properties': {
                        'title': sheetName
                    }
                }
            }
        ],
    }

    request = gc.service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_spreadsheet_request_body)
    response = request.execute()

