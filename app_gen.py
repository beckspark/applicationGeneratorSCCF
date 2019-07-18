import os
import pandas as pd
import glob
import shutil
import subprocess
import sys


def get_dict():
    global df
    df = pd.read_csv(os.path.join('processing','client.csv'))
    df = df.iloc[0].to_dict()


def continue_on():
    cont = 'no'
    while cont != 'y':
        cont = input('Enter \'y\' to continue, or enter \'e\' to exit: ')
        if cont == 'e':
            sys.exit()


def exit_now():
    cont = 'no'
    while cont != 'e':
        cont = input('Enter \'e\' to exit: ')
    sys.exit()


def dropbox_setup():
    global db_dir, pen_dir
    db_dir = os.path.join("\\\\10.1.10.201\\Dropbox\\Services")
    if not os.path.exists(db_dir):
        db_dir = (os.path.join("%s","Dropbox\\Services" % os.path.expanduser('~')))
        if not os.path.exists(db_dir):
            print('The Dropbox folder doesn\'t seem to be properly set up! Script will now exit.')
            exit_now()
    pen_dir = os.path.join(db_dir,"Pending_Applications")


def find_app():
    global app_file, app_abs
    if not glob.glob(os.path.join(pen_dir,"*.pdf")):
        print('Could not find an application to process!')
        exit_now()
    else:
        for file in os.listdir(pen_dir):
            if file.endswith(".pdf") or file.endswith(".PDF"):
                cont = 'no'
                while cont != 'y':
                    cont = input('Found the application '+file+'. Process? \'y\' to continue: ')
                    if cont == 'n':
                        exit_now()
                app_file = os.path.join(file)
                app_abs = os.path.join(pen_dir,app_file)


def proc_pdf():
    global proc_dir, app_abs
    proc_dir = os.path.join(pen_dir,"processing")
    shutil.copy(app_abs,proc_dir)
    os.remove(app_abs)
    app_abs = os.path.join(proc_dir,app_file)


def pdf_open():
    global sum_open
    sumatra = r'C:\\Program Files\\SumatraPDF\\SumatraPDF.exe'
    sum_open = subprocess.Popen([sumatra, app_abs], shell=True)
    print('The application will now open in a new window. Refer to it to answer quesitons about the client. You do not need to exit the application, so please DO NOT.')


if __name__=='__main__':
    try:
        print('Welcome to the Application Processing App for Senior Charity Care Foundation!')
        dropbox_setup()
        find_app()
        proc_pdf()
        pdf_open()
        client_data.get_data()
        get_dict()
        if df['approval'] == 'Approval':
            if 'Dental' in df['service']:
                approval.letter_contents('Dental')
                approval()
                dentimax()
            if 'Vision' in df['service']:
                df.to_csv(os.path.join('processing','v_client.csv'), index = False)
            if 'Hearing' in df['service']:
                df.to_csv(os.path.join('processing','h_client.csv'), index = False)
        elif df['approval'] == 'Denial':
            denial.letter()
        data_write()
        org_files()
        print('Application processed!')
        exit_now()
    except e as Exception:
        print(e)
        exit_now()
