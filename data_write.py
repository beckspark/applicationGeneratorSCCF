import pandas as pd
import os
from __init__ import *
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def update_csv_data(d):
    df_csv = os.path.join('processing',(d+'.csv'))
    df = pd.read_csv(df_csv)
    df = df.drop([0])
    df.to_csv(df_csv)


def drive_add_data(data):
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
    gc = gspread.authorize(credentials)
    wks = gc.open("sccf").worksheet(data)
    wks.append_row(df)


def get_dict(data):
    global df
    df = pd.read_csv(os.path.join('processing',data)).to_dict(orient='records')
    df = list(df[0].values())


def write_data():
    get_dict('client.csv')
    drive_add_data('client')
    get_dict('demo.csv')
    drive_add_data('demographics')
    update_csv_data('client')
    update_csv_data('demo')


if __name__ == '__main__':
    write_data()
