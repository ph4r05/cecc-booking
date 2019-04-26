import xlsxwriter
import itertools
import json
import collections
import sys
import math

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


class Renderer(object):
    def __init__(self, worksheet, offset, key, rooms, double_col=True):
        self.worksheet = worksheet
        self.room_size = max([x['beds'] for x in rooms])
        self.double_col = double_col
        self.title_span = 2*4+1 if self.double_col else 4
        self.offset = offset
        self.key = key
        self.rooms = rooms
        self.room_offset = None
        self.offset_final = None
        self.edits = []

    def draw_type(self):
        self.worksheet.merge_range(self.offset, 0, self.offset, self.title_span - 1, get_room_label(self.key), room_type_format)
        self.offset += 1

    def draw_headings(self):
        self.worksheet.write_string(self.offset, 0, 'Room ID')
        self.worksheet.merge_range(self.offset, 1, self.offset, 3, 'Name Surname', names_format)
        self.offset += 1

        if not self.double_col:
            return

        self.offset -= 1
        self.worksheet.write_string(self.offset, 5, 'Room ID')
        self.worksheet.merge_range(self.offset, 6, self.offset, 8, 'Name Surname', names_format)
        self.offset += 1

    def begin_rooms(self):
        self.room_offset = self.offset
        nroom_rows = len(self.rooms)
        if self.double_col:
            nroom_rows = math.ceil(nroom_rows/2.)

        self.offset_final = self.room_offset + (self.room_size + 1) * nroom_rows

    def get_room_start(self, idx):
        if not self.double_col:
            return 0, self.room_offset + idx * (self.room_size + 1)

        x_coord = 0 if (idx & 1) == 0 else 5
        y_coord = self.room_offset + idx//2 * (self.room_size + 1)
        return x_coord, y_coord

    def draw_room(self, idx):
        coords = self.get_room_start(idx)
        room = self.rooms[idx]
        room_id = '%s' % get_room_id(room['id'])
        num_ppl = len(room['people'])
        num_beds = room['beds']

        # room id
        self.worksheet.merge_range(coords[1], coords[0], coords[1] + self.room_size - 1, coords[0], room_id, room_id_format)

        offset = coords[1]
        for cbed in range(num_beds):
            cstr = 'Free' if cbed >= num_ppl else (room['people'][cbed] if show_names else 'Taken')
            cstr = '%d. %s' % (cbed + 1, cstr)
            cstyle = name_format_empty if cbed >= num_ppl else name_format

            self.edits.append({
                'row': offset, 'col': coords[0] + 1, 'lcol': coords[0] + 1,
                'body': cstr, 'taken': cbed < num_ppl
            })
            self.worksheet.merge_range(offset, coords[0] + 1, offset, coords[0] + 3, cstr, cstyle)
            offset += 1
        self.offset = offset


worksheet.merge_range('A1:I1', 'CECC 2019 Bookings', merge_format)
# worksheet.set_column(0, 1, 10)

offset = 1
booking_data.sort(key=sorter)
booking_data_disp = booking_data if show_lectors else filter(filterer, booking_data)
edits = []

for k, g in itertools.groupby(booking_data_disp, key=lambda x: x['type']):
    g = list(g)
    renderer = Renderer(worksheet, offset, k, g, True)
    renderer.draw_type()
    renderer.draw_headings()
    renderer.begin_rooms()
    for i in range(len(g)):
        renderer.draw_room(i)

    offset = renderer.offset + 1
    edits += renderer.edits

workbook.close()
# print(json.dumps(edits, indent=2))

# Google Docs code
# https://developers.google.com/sheets/api/samples/writing
# https://developers.google.com/sheets/api/guides/values#writing
# https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#PasteDataRequest
# https://developers.google.com/sheets/api/guides/batchupdate
# https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/sheets_v4.spreadsheets.html#batchUpdate

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

    for i in range(ed['col'], ed['lcol']+1):
        edits_map[ed['row']][i] = ed['taken']

first_edit_row = min(edits, key=lambda x: x['row'])
last_edit_row = max(edits, key=lambda x: x['row'])
first_edit_col = min(edits, key=lambda x: x['col'])
last_edit_col = max(edits, key=lambda x: x['col'])
frow, fcol = first_edit_row['row'], first_edit_col['col']
lrow, lcol = last_edit_row['row'], last_edit_col['col']

cell_request = {
    # 'range': {
    #     'sheetId': None,
    #     'startRowIndex': frow,
    #     'endRowIndex': lrow,
    #     'startColumnIndex': fcol+1,
    #     'endColumnIndex': lcol+1,
    # },
    'start': {
      'sheetId': None,
      'rowIndex': frow,
      'columnIndex': fcol
    },
    'rows': [],
    'fields': 'userEnteredFormat.backgroundColor'
}


# print(json.dumps(edits_map, indent=2))
for crow in range(frow, lrow + 1):
    crow_data = [None] * (lcol - fcol + 1)

    for cidx, ccol in enumerate(range(fcol, lcol + 1)):
        crow_data[cidx] = {}
        if edits_map[crow][ccol] is not None:
            crow_data[cidx] = {}
            crow_data[cidx]['userEnteredFormat'] = {
                'backgroundColor': color_taken if edits_map[crow][ccol] else color_free
            }
            #'stringValue': str(cidx)

    cell_request['rows'].append({'values': crow_data})
print(json.dumps(cell_request, indent=2))

body = {
    'valueInputOption': 'USER_ENTERED',
    'data': edit_data,
}

# Get sheet ID
spreadsheet_info = sheet.get(spreadsheetId=SPREADSHEET_ID, ranges=[]).execute()
cell_request['start']['sheetId'] = spreadsheet_info['sheets'][0]['properties']['sheetId']

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
