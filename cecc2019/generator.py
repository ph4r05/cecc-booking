import xlsxwriter
import itertools
import json
import collections
import sys

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
# SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1mfH5Ql6QbWK95xVtbxLmVbBLM84PRRmHPk3O74cLswA'
show_names = False
show_lectors = False

booking_data = json.load(open('bookings.json'))

stars = [
    'MiaKhalifa',
    'LittleCaprice',
    'MiaMalkova',
    'ChaseyLane',
    'RileyReid',
    'MarshaMay',
    'BrandyLove',
    'DillionHarper',
    'SashaGrey',
    'KatanaKombat',
    'AbellaDanger',
    'KarmaRx',
    'JadeKush',
    'AdrianaChechik',
    'SarahBanks',
    'LanaRhodes',
    'QuinnWilde',
    'LeahGotti',
    'ArianaMarie',
    'ElsaJean',
    'BrettRosi',
    'DaniDaniels',
]


def get_room_label(x):
    if x == '3':
        return 'Wheelchair accessible'
    elif x == '4':
        return '4-bed'
    elif x == 'm2':
        return '2-bed duplex'
    elif x == 'm5':
        return '5-bed duplex'
    else:
        raise ValueError('Unknown type: ' + x)


def get_room_id(rid):
    if False and len(stars) > rid:
        return stars[rid]
    else:
        return rid


def col_txt(idx):
    return chr(ord('A') + idx)


def coords_txt(frow, fcol, lrow=None, lcol=None):
    base = '%s%s' % (col_txt(fcol), frow)
    if lrow is not None:
        base += ':%s%s' % (col_txt(lcol), lrow)
    return base


filterer = lambda x: x['type'] in ['3', '4']
sorter = lambda x: (x['type'], x['id'])

workbook = xlsxwriter.Workbook('cecc2019.xlsx')
worksheet = workbook.add_worksheet()

merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_size': 14,
})

room_type_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
})

room_id_format = workbook.add_format({
    'bold': 0,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter',
})

names_format = workbook.add_format({
    'bold': 0,
    'border': 0,
    'align': 'left',
    'valign': 'vcenter',
})

name_format = workbook.add_format({
    'bold': 0,
    'border': 0,
    'align': 'left',
    'valign': 'vcenter',
    'fg_color': 'red',
})

name_format_empty = workbook.add_format({
    'bold': 0,
    'border': 0,
    'align': 'left',
    'valign': 'vcenter',
    'fg_color': 'green',
})

color_taken = {
    'red': 1,
    'green': 0,
    'blue': 0,
    'alpha': 1
}

color_free = {
    'red': 0,
    'green': 1,
    'blue': 0,
    'alpha': 1
}

worksheet.merge_range('A1:D1', 'CECC 2019 Bookings', merge_format)
# worksheet.set_column(0, 1, 10)

offset = 0
booking_data.sort(key=sorter)
booking_data_disp = booking_data if show_lectors else filter(filterer, booking_data)
edits = []

for k, g in itertools.groupby(booking_data_disp, key=lambda x: x['type']):
    offset += 2
    worksheet.merge_range(offset, 0, offset, 3, get_room_label(k), room_type_format)
    offset += 1
    worksheet.write_string(offset, 0, 'Room ID')
    worksheet.merge_range(offset, 1, offset, 3, 'Name Surname', names_format)

    for room in g:
        offset += 1
        # worksheet.write_string(offset, 0, '%s' % get_room_id(room['id']))
        worksheet.merge_range(offset, 0, offset + room['beds'] - 1, 0, '%s' % get_room_id(room['id']), room_id_format)
        num_ppl = len(room['people'])

        for cbed in range(room['beds']):
            cstr = 'Free' if cbed >= num_ppl else (room['people'][cbed] if show_names else 'Taken')
            cstr = '%d. %s' % (cbed + 1, cstr)
            cstyle = name_format_empty if cbed >= num_ppl else name_format

            edits.append({
                'row': offset, 'col': 1, 'body': cstr, 'taken': cbed < num_ppl
            })
            worksheet.merge_range(offset, 1, offset, 3, cstr, cstyle)
            offset += 1

workbook.close()
print(json.dumps(edits, indent=2))

# Google Docs code
# https://developers.google.com/sheets/api/samples/writing
# https://developers.google.com/sheets/api/guides/values#writing
# https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#PasteDataRequest
# https://developers.google.com/sheets/api/guides/batchupdate
# https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/sheets_v4.spreadsheets.html#batchUpdate
#
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server()

    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)


service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

edit_data = []
edits_map = collections.defaultdict(lambda: collections.defaultdict(lambda: None))
for ed in edits:
    edit_data.append({
        'range': 'Sheet1!%s' % coords_txt(ed['row'] + 1, ed['col']),
        'values': [[ed['body']]]
    })

    edits_map[ed['row']][ed['col']] = ed['taken']

first_edit = min(edits, key=lambda x: x['row'])
last_edit = max(edits, key=lambda x: x['row'])
frow, fcol = first_edit['row'], first_edit['col']
lrow, lcol = last_edit['row'], last_edit['col']

# Get sheet ID
spreadsheet_info = sheet.get(spreadsheetId=SPREADSHEET_ID, ranges=[]).execute()

cell_request = {
    # 'range': {
    #     'sheetId': spreadsheet_info['sheets'][0]['properties']['sheetId'],
    #     'startRowIndex': frow,
    #     'endRowIndex': lrow,
    #     'startColumnIndex': fcol+1,
    #     'endColumnIndex': lcol+1,
    # },
    'start': {
      'sheetId': spreadsheet_info['sheets'][0]['properties']['sheetId'],
      'rowIndex': frow,
      'columnIndex': fcol
    },
    'rows': [],
    'fields': 'userEnteredFormat'
}


# print(json.dumps(edits_map, indent=2))
for crow in range(frow, lrow + 1):
    crow_data = [None] * (lcol - fcol + 1)
    changed = False

    for cidx, ccol in enumerate(range(fcol, lcol + 1)):
        # if edits_map[crow][ccol] is None:
        #     continue

        changed = True
        crow_data[cidx] = {}
        if edits_map[crow][ccol] is not None:
            crow_data[cidx]['userEnteredFormat'] = {
                'backgroundColor': color_taken if edits_map[crow][ccol] else color_free
            }
            #'stringValue': str(cidx)

    cell_request['rows'].append({'values': crow_data} if changed else {'values': {}})

body = {
    'valueInputOption': 'USER_ENTERED',
    'data': edit_data,
}

# print(json.dumps(body, indent=2))
result = service.spreadsheets().values().batchUpdate(
    spreadsheetId=SPREADSHEET_ID, body=body).execute()
print('{0} cells updated.'.format(result.get('updatedCells')))

print(json.dumps(cell_request))#, indent=2))
result = service.spreadsheets().batchUpdate(
    spreadsheetId=SPREADSHEET_ID, body={
        'requests': [
            {'updateCells': cell_request},
        ]
    }).execute()
print('{0} cells updated.'.format(result.get('updatedCells')))

# result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
#                             range=RANGE_NAME).execute()
# print(result)

print(json.dumps(booking_data, indent=2))
