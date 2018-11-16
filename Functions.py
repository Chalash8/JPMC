from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import pygsheets
import datetime
from datetime import date
from decimal import *

SCOPES = 'https://www.googleapis.com/auth/drive'

def get_drive_service():
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('drive', 'v3', http=creds.authorize(Http()))
    return service


def get_timesheet_id(week_end):
    results = get_drive_service().files().list(q="'0Bzzqy5DUeqkyZ1o5azktY3ZscnM' in parents", fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        for item in items:
            if item['name'] == week_end:
                return item['id']

def get_spreadsheet_values(ss_id, wks_title):
    c = pygsheets.authorize()
    sh = c.open_by_key(ss_id)
    wks = sh.worksheet_by_title(wks_title)
    cell_matrix = wks.get_all_values(returnas='matrix')
    return cell_matrix

def convert_to_date(date_str):
    if len(date_str) == 12:
        y = int(date_str[slice(2,6)])
        m = int(date_str[slice(7,9)])
        d = int(date_str[slice(10,12)])
        return datetime.date(y, m, d)
    elif len(date_str) == 10:
        y = int(date_str[slice(6,10)])
        m = int(date_str[slice(0,2)])
        d = int(date_str[slice(3,5)])
        return datetime.date(y, m, d)
    else:
        print('Bad date_str in convert_to_date')


def write_to_spreadsheet(ss_id,cell_matrix,invoice_num):
    c = pygsheets.authorize()
    sh = c.open_by_key(ss_id)
    wks = sh.sheet1
    wks.clear(start='A2')
    row = 1
    for i in cell_matrix:
        ot=0
        if Decimal(i[9]) > 40:
            ot=40-Decimal(i[9])
            st=40
        else:
            st=i[9]
        writeable_row = [i[4],i[1],i[2],i[0],i[14],i[12],st,ot,i[9]]
        if i[1] == 'JPCH-GD' and i[14] != '351393':
            print('ERROR: '+ i[1] + ' is not a Chicago USCIS worker')
        if i[1] == 'JPVT-GD' and i[14] != '028065':
            print('ERROR: '+ i[1] + ' is not a Vermont CH worker')
        wks.insert_rows(row,values=writeable_row)
        row+=1

        wks.update_value('J2',invoice_num)
        wks.update_value('K2', str(date.today()))
    print (cell_matrix[0][1] + " has been written")

def run_script():
    c = pygsheets.authorize()


def get_most_recent_file_id(folder_id):
    results = get_drive_service().files().list(q="'{}' in parents".format(folder_id), fields="nextPageToken, files(createdTime, id, name)").execute()
    items = results.get('files', [])
    items.sort
    if not items:
        print('No files found.')
    else:
        return items[0]['id']


def copy_old_ss(ss_id,week_end_str):
    service = get_drive_service()
    file = service.files().get(fileId=ss_id).execute()
    if file['name'] == week_end_str or file['name'] ==  week_end_str+" CH":
        print('File aready exists:')
        print(file)
        return ss_id
    elif "CH" in file['name']:
        copied_file = {'name': week_end_str+" CH"}
        return service.files().copy(fileId=ss_id, body=copied_file).execute()['id']
    else:
        copied_file = {'name': week_end_str}
        return service.files().copy(fileId=ss_id, body=copied_file).execute()['id']

def get_invoice_num(we_date):
    service = get_drive_service()
    we_date = we_date - datetime.timedelta(days=7)
    we_date_str = 'WE' + str(we_date)
    ss_id = get_timesheet_id(we_date_str)
    return raw_input('Enter the invoice number for ' + ss_id + ': ')


def temp():
    if ss_id == None:
        return raw_input('Enter the invoice number for Chicago USCIS: ') 
    else:
        c = pygsheets.authorize()
        sh = c.open_by_key(ss_id)
        wks = sh.sheet1
        cell = wks.cell('j2')
        invoice_num = int(cell.value)

        print(invoice_num)
        return invoice_num + 1
