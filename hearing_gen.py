import pandas as pd
import os
from datetime import datetime
from string import Template
from approval import letter_word
from org_files import client_directory, client_files
from app_gen import exit_now


def rm_client_df():
    df = df.drop[0]
    df.to_csv(h_dict, index = False)

def get_h_dict():
    global df, df_client, h_dict
    h_dict = os.path.join('processing','h_client.csv')
    print(h_dict)
    if os.path.isfile(h_dict):
        df = pd.read_csv(h_dict)
        df_client = df.iloc[0].to_dict()
    else:
        print('No Hearing Application found to process!')
        exit_now()


def h_letter():
    global letter
    approval_template = open(os.path.join('processing','hearing_approval.txt'))
    src = Template(approval_template.read())
    h = {'date':df['approval_date'],'fname':df['fname'],'lname':df['lname'],'housing':df['housing'],'address':df['address'],'apt':df['unit'],'city':df['city'],'zipCode':df['zip_code'],'phone':df['phone'],'percent':df['percent'],'aidCost':df['aid_cost'],'aidCost2':df['aid_cost_2'],'yourName':yourName,'title':title}
    letter = src.substitute(h)


if __name__=='__main__':
    print('Welcome to the application generator for Hearing Clients at SCCF!: ')
    year = datetime.strftime(datetime.today(), "%Y")
    yourName = input('What is your full name?: ')
    title = input('What is your title at SCCF?: ')
    get_h_dict()
    client_name = df_client['lname'] + "_" + df_client['fname']
    h_letter()
    letter_word('Hearing', letter, df_client['lname'], df_client['fname'], datetime.strftime(datetime.today(),"%m.%d.%Y"))
    client_directory(client_name)
    client_files(client_name, os.path.join(os.path.expanduser('~'),'sccf','Clients',client_name,year))
    rm_client_df()
