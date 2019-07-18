import os
import shutil
import glob
import pandas as pd
from datetime import datetime


year = datetime.strftime(datetime.today(), "%Y")


def get_dict():
    global df
    df = pd.read_csv(os.path.join('processing','client.csv')).to_dict(orient='records')
    df = df[0]


def client_directory(c):
    global client_dir
    client_dir = os.path.join(os.path.expanduser('~'),'sccf','Clients',c,year)
    if not os.path.exists(client_dir):
        os.makedirs(client_dir)


def client_files(c, d):
    file_search = os.path.join('processing',(c + "*"))
    for files in glob.glob(file_search):
        if not os.path.isdir(files):
            shutil.copy(files,d)
            os.remove(files)


if __name__=="__main__":
    get_dict()
    client_directory(df['lname'], df['fname'])
    client_files(df['lname'], df['fname'], client_dir)
