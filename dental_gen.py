import pandas as pd
import os
from datetime import datetime
from string import Template
from approval import letter_word
from org_files import client_directory, client_files
from app_gen import exit_now


def get_d_dict():
    global df
    df = pd.read_csv(os.path.join('processing','client.csv'))
    df = df.iloc[0].to_dict()


def d_letter():
    global letter
    approval_template = open(os.path.join('processing','dental_approval.txt'))
    src = Template(approval_template.read())
    d = {'date':df['approval_date'],'fname':df['fname'],'lname':df['lname'],'housing':df['housing'],'address':df['address'],'apt':df['unit'],'city':df['city'],'zipCode':df['zip_code'],'phone':df['phone'],'hourly':df['hourly'],'denture':df['denture'],'dentures':df['dentures'],'flipper':df['partial'],'valplast':df['valplast'],'percent':df['percent'],'yourName':yourName,'title':title}
    letter = src.substitute(d)


if __name__=='__main__':
    print('Welcome to the application generator for Dental Clients at SCCF!: ')
    year = datetime.strftime(datetime.today(), "%Y")
    yourName = input('What is your full name?: ')
    title = input('What is your title at SCCF?: ')
    get_d_dict()
    client_name = df['lname'] + "_" + df['fname']
    d_letter()
    letter_word('Dental', letter, df['lname'], df['fname'], datetime.strftime(datetime.today(),"%m.%d.%Y"))
    client_directory(client_name)
    client_files(client_name, os.path.join(os.path.expanduser('~'),'sccf','Clients',client_name,year))
