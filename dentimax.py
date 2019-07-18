import pandas as pd
import os
import win32com.client as win32
import pywinauto
import shutil
import subprocess
import pyperclip
from string import Template
import pywinauto
from pywinauto import *
from time import sleep
from app_gen import *
from client_data import *


def get_dict():
    global df
    df = pd.read_csv(os.path.join('processing','client.csv'))
    df = df.iloc[0].to_dict()


def dmax_init():
    print('\nDentiMax will start in a separate window. Please make sure you are completely logged in, and keep in mind you must NOT touch the keyboard or mouse while we\'re working in DentiMax. Continue only once you\'re ready.\n')
    continue_on()
    dmax = 'dentimax.rdp'
    dmax_run = subprocess.Popen(['mstsc',dmax], shell = True)


def dmax_define():
    global app, h, dm_win
    app = application.Application()
    h = pywinauto.findwindows.find_windows(title_re=".*GATEWAY1",class_name="RAIL_WINDOW")
    if len(h) > 1:
        app.connect(handle=h[1])
    elif len(h) <= 1:
        app.connect(handle=h[0])
    dm_win = app.top_window()


def tab_sleep(n):
    for _ in range(n):
        dm_win.type_keys('{TAB}')
        sleep(.1)


def dmax_client_add():
    print('We will now add this client to DentiMax. Please do NOT touch the mouse or keyboard while this is being done!\n')
    continue_on()
    dm_win.type_keys('{F4}') # Returns to home screen in DentiMax just in case
    sleep(3)
    dm_win.type_keys('%p')
    sleep(3)
    dm_win.type_keys('{F8}')
    sleep(3)
    tab_sleep(2)
    dm_win.type_keys(df['lname'].item(), with_spaces=True)
    tab_sleep(1)
    dm_win.type_keys(df['fname'].item(), with_spaces=True)
    tab_sleep(4)
    dm_win.type_keys(df['housing'].item(), with_spaces=True)
    tab_sleep(1)
    dm_win.type_keys(df['unit'].item())
    tab_sleep(1)
    dm_win.type_keys(df['zip_code'].item())
    tab_sleep(1)
    dm_win.type_keys(df['phone'].item())
    tab_sleep(6)
    pyperclip.copy(df['dob'].item())
    dm_win.type_keys('^v')
    tab_sleep(10)
    dm_win.type_keys('Application approved and processed - '+yourName+' '+date_save+'. Hourly rate is $'+str(df['hourly'].item()), with_spaces=True)
    dm_win.type_keys('{F3}')
    print('Done! Please double-check that they were added correctly.')
    continue_on()


def dmax_lookup_client():
    print('\n We will now take a look at whether or not this client is in DentiMax. \n')
    continue_on()
    dm_win.type_keys('{F4}') # Returns to home screen in DentiMax just in case
    sleep(3)
    dm_win.type_keys('%p')
    sleep(3)
    tab_sleep(3)
    dm_win.type_keys(df['lname'].item(), with_spaces=True)
    tab_sleep(1)


if __name__=='__main__':
    proc_dir = "processing"
    date_save = datetime.strftime(datetime.today(), "%m.%d.%Y")
    yourName = 'Stephen Beck'
    title = 'AmeriCorps Program Director'
    get_client_dict()
    dmax_define()
    dmax_client_add()
