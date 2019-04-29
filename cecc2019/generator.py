import xlsxwriter
import itertools
import json
import collections
import sys
import re
import copy
import math
import logging
import argparse
import coloredlogs

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
ADMINSPREADSHEET_ID = '1dHCTgInqqmSyLMS64eGgWucA2PzNtsyfei-yUew0EfI'
show_names = True
show_lectors = True

logger = logging.getLogger(__name__)
coloredlogs.CHROOT_FILES = []
coloredlogs.install(level=logging.WARNING, use_chroot=False)

parser = argparse.ArgumentParser(description="CECC 2019 accommodation booking script")
parser.add_argument("--load", dest="load", default=False, action="store_const", const=True, help="Load bookings from the Admin spreadsheet",)
parser.add_argument("--no-sync", dest="no_sync", default=False, action="store_const", const=True, help="No Google sync",)
args = parser.parse_args()


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


def strip_index(txt):
    m = re.match('^\s*\d+\.\s*(.*)$', txt)
    return m.group(1) if m else txt


def is_bed_free(txt):
    if 'free' in txt.lower() or len(txt.strip()) == 0:
        return True
    if re.match('^\s*\d+\.\s*$', txt) is not None:
        return True
    return False


filterer = lambda x: x['type'] in ['3', '4']
sorter = lambda x: (x['type'], x['id'])
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
    def __init__(self, worksheet, offset, key, rooms, double_col=True, booking=None):
        self.worksheet = worksheet
        self.room_size = max([x['beds'] for x in rooms])
        self.double_col = double_col
        self.title_span = 2*4+1 if self.double_col else 4
        self.offset = offset
        self.key = key
        self.rooms = rooms
        self.room_offset = None
        self.offset_final = None
        self.booking = booking
        self.do_conditional_formatting = False
        self.edits = []

    def draw_type(self):
        self.worksheet.merge_range(self.offset, 0, self.offset, self.title_span - 1, get_room_label(self.key), self.booking.room_type_format)
        self.offset += 1

    def draw_headings(self):
        self.worksheet.write_string(self.offset, 0, 'Room ID')
        self.worksheet.merge_range(self.offset, 1, self.offset, 3, 'Name Surname', self.booking.names_format)
        self.offset += 1

        if not self.double_col:
            return

        self.offset -= 1
        self.worksheet.write_string(self.offset, 5, 'Room ID')
        self.worksheet.merge_range(self.offset, 6, self.offset, 8, 'Name Surname', self.booking.names_format)
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
        num_ppl = min(len(room['people']), sum([1 for x in room['people'] if x is not None]))
        num_beds = room['beds']

        # room id
        self.worksheet.merge_range(coords[1], coords[0], coords[1] + self.room_size - 1, coords[0], room_id, self.booking.room_id_format)

        offset = coords[1]
        for cbed in range(num_beds):
            cstr = 'Free' if cbed >= num_ppl else (room['people'][cbed] if self.booking.show_names else 'Taken')
            cstr = '%d. %s' % (cbed + 1, cstr)
            cstyle = self.booking.name_format_empty if cbed >= num_ppl else self.booking.name_format

            self.edits.append({
                'row': offset, 'col': coords[0] + 1, 'lcol': coords[0] + 1,
                'body': cstr, 'taken': cbed < num_ppl,
                'num_ppl': num_ppl,
                'cbed': cbed,
                'person': room['people'][cbed] if cbed < num_ppl else None,
                'room_idx': idx,
                'room_key': self.key,
                'room_id': room['id'],
            })

            self.worksheet.merge_range(offset, coords[0] + 1, offset, coords[0] + 3, cstr, cstyle if not self.do_conditional_formatting else None)

            # https://xlsxwriter.readthedocs.io/example_conditional_format.html
            if self.do_conditional_formatting:
                self.worksheet.conditional_format(
                    offset, coords[0] + 1, offset, coords[0] + 1, {
                        'type': 'text',
                        'criteria': 'containsText',
                        'value': 'Free',
                        'format': self.booking.name_format_empty_bg
                    }
                )

                self.worksheet.conditional_format(
                    offset, coords[0] + 1, offset, coords[0] + 1, {
                        'type': 'text',
                        'criteria': 'notContains',
                        'value': 'Free',
                        'format': self.booking.name_format_bg
                    }
                )

            offset += 1
        self.offset = offset


class Bookings:
    def __init__(self):
        # https://xlsxwriter.readthedocs.io/
        # https://xlsxwriter.readthedocs.io/examples.html
        self.workbook = None
        self.worksheet = None
        self.args = None
        self.creds = None
        self.show_names = show_names
        self.do_sync_to_admin = False
        self.do_conditional_formatting = True
        self.booking_data = None
        self.edits = None
        self.merge_format = None
        self.room_type_format = None
        self.room_id_format = None
        self.names_format = None
        self.name_format = None
        self.name_format_empty = None

    def work(self, args):
        self.args = args
        self.booking_data = json.load(open('bookings.json'))
        self.booking_data.sort(key=sorter)

        self.show_names = self.do_sync_to_admin
        self.gen()

        if self.args.no_sync:
            return

        self.load_creds()

        # Sync current JSON to admin spreadsheet.
        # Warning: clears current modifications.
        if self.do_sync_to_admin:
            self.sync_write(ADMINSPREADSHEET_ID)

        # Reads values from the admin spreadsheet.
        self.sync_read()

        # Re-generate with the anonymous names, fetched from admin
        # Re-generate excel sheet so we have anonymous ready for upload, with new
        # updated data.
        self.show_names = False
        self.gen()
        self.sync_write()

    def gen_styles(self):
        self.merge_format = self.workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14,
        })

        self.room_type_format = self.workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
        })

        self.room_id_format = self.workbook.add_format({
            'bold': 0,
            'border': 0,
            'align': 'center',
            'valign': 'vcenter',
        })

        self.names_format = self.workbook.add_format({
            'bold': 0,
            'border': 0,
            'align': 'left',
            'valign': 'vcenter',
        })

        self.name_format = self.workbook.add_format({
            'bold': 0,
            'border': 0,
            'align': 'left',
            'valign': 'vcenter',
            'fg_color': 'red',
        })

        self.name_format_empty = self.workbook.add_format({
            'bold': 0,
            'border': 0,
            'align': 'left',
            'valign': 'vcenter',
            'fg_color': 'green',
        })

        self.name_format_bg = self.workbook.add_format({
            'bg_color': 'red'
        })

        self.name_format_empty_bg = self.workbook.add_format({
            'bg_color': 'green'
        })

    def gen(self):
        self.workbook = xlsxwriter.Workbook('cecc2019.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.gen_styles()
        self.worksheet.merge_range('A1:I1', 'CECC 2019 Bookings', self.merge_format)
        # self.worksheet.set_column(0, 1, 10)

        offset = 1
        booking_data_disp = self.booking_data if show_lectors else filter(filterer, self.booking_data)
        self.edits = []

        for k, g in itertools.groupby(booking_data_disp, key=lambda x: x['type']):
            g = list(g)
            renderer = Renderer(self.worksheet, offset, k, g, True, booking=self)
            renderer.do_conditional_formatting = self.do_conditional_formatting
            renderer.draw_type()
            renderer.draw_headings()
            renderer.begin_rooms()
            for i in range(len(g)):
                renderer.draw_room(i)

            offset = renderer.offset + 1
            self.edits += renderer.edits

        self.workbook.close()
        # print(json.dumps(edits, indent=2))

    def load_creds(self):
        # Google Docs code
        # https://developers.google.com/sheets/api/samples/writing
        # https://developers.google.com/sheets/api/guides/values#writing
        # https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#PasteDataRequest
        # https://developers.google.com/sheets/api/guides/batchupdate
        # https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/sheets_v4.spreadsheets.html#batchUpdate

        self.creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)

        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                self.creds = flow.run_local_server()

            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

    def sync_write(self, spreadsheet_id=SPREADSHEET_ID):
        service = build('sheets', 'v4', credentials=self.creds)
        sheet = service.spreadsheets()

        edit_data = []
        edits_map = collections.defaultdict(lambda: collections.defaultdict(lambda: None))
        for ed in self.edits:
            edit_data.append({
                'range': 'Sheet1!%s' % coords_txt(ed['row'] + 1, ed['col']),
                'values': [[ed['body']]]
            })

            for i in range(ed['col'], ed['lcol']+1):
                edits_map[ed['row']][i] = ed['taken']

        first_edit_row = min(self.edits, key=lambda x: x['row'])
        last_edit_row = max(self.edits, key=lambda x: x['row'])
        first_edit_col = min(self.edits, key=lambda x: x['col'])
        last_edit_col = max(self.edits, key=lambda x: x['col'])
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
        spreadsheet_info = sheet.get(spreadsheetId=spreadsheet_id, ranges=[]).execute()
        cell_request['start']['sheetId'] = spreadsheet_info['sheets'][0]['properties']['sheetId']

        # Sheet text data udpate
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id, body=body).execute()
        print('{0} cells updated.'.format(result.get('updatedCells')))

        # Sheet formatting update
        # Update is not needed if we have conditional formatting.
        if self.do_conditional_formatting:
            return

        print(json.dumps(cell_request))#, indent=2))
        result = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={
                'requests': [
                    {'updateCells': cell_request},
                ]
            }).execute()
        print('{0} cells updated.'.format(result.get('updatedCells')))
        print(json.dumps(self.booking_data, indent=2))

    def sync_read(self):
        service = build('sheets', 'v4', credentials=self.creds)
        sheet = service.spreadsheets()

        booking_map = collections.defaultdict(lambda: None)  # room id -> room
        edits_map = collections.defaultdict(lambda: collections.defaultdict(lambda: None))
        for ed in self.edits:
            for i in range(ed['col'], ed['lcol'] + 1):
                edits_map[ed['row']][i] = ed

        for bo in self.booking_data:
            room_copy = copy.deepcopy(bo)
            room_copy['people'] = []  # clear as we are going to append
            booking_map[bo['id']] = room_copy

        result = sheet.values().get(spreadsheetId=ADMINSPREADSHEET_ID,
                                    range='!A1:I').execute()

        if result['majorDimension'] != 'ROWS':
            raise ValueError('Not supported')

        rows = result['values']
        for irow, row in enumerate(rows):
            for icol, col in enumerate(row):
                if edits_map[irow][icol] is None:
                    continue

                rec = edits_map[irow][icol]
                room_id = rec['room_id']

                if booking_map[room_id] is None:
                    logger.warning('Room ID %s not found in booking data' % room_id)
                    continue

                room = booking_map[room_id]
                is_free = is_bed_free(col)
                cname = strip_index(col)
                room['people'].append((rec['cbed'], None if is_free else cname))

        for room_id in booking_map:
            room = booking_map[room_id]
            room['people'] = [x[1] for x in sorted(room['people'], key=lambda x: x[0])]

        rooms = [x[2] for x in sorted([(booking_map[room_id]['type'], room_id, booking_map[room_id]) for room_id in booking_map])]
        self.booking_data = rooms
        json.dump(rooms, open('bookings_fetched.json', 'w+'), indent=2)


def main(args):
    booking = Bookings()
    booking.work(args)
    # loop = asyncio.get_event_loop()
    # loop.run_until_complete(amain(args))
    # loop.close()


if __name__ == "__main__":
    main(args)
