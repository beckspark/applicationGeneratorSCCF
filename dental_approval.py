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
from __init__ import *
from client_data import *


def get_client_dict():
    global df
    df = pd.read_csv(os.path.join('processing','client.csv'))
    print(df['lname'].item())


def dental_letter_contents():
    global approval_letter
    approval_template = open('dental_approval.txt')
    src = Template(approval_template.read())
    d = {'date':df['approval_date'].item(),'fname':df['fname'].item(),'lname':df['lname'].item(),'housing':df['housing'].item(),'address':df['address'].item(),'apt':df['unit'].item(),'city':df['city'].item(),'zipCode':df['zip_code'].item(),'phone':df['phone'].item(),'hourly':df['hourly'].item(),'denture':df['denture'].item(),'dentures':df['dentures'].item(),'flipper':df['partial'].item(),'valplast':df['valplast'].item(),'percent':df['percent'].item(),'yourName':yourName,'title':title}
    approval_letter = src.substitute(d)


def dental_letter_word():
    letterhead = 'application_template.docx'
    word_save_name = df['lname'].item() + "_" + df['fname'].item() + "_Dental_Approval_Letter_" + date_save + ".docx"
    word_save = os.path.join(word_save_name)
    shutil.copy(letterhead, os.path.join(proc_dir, word_save))
    word = win32.DispatchEx('Word.Application')
    doc = word.Documents.Open(os.path.join(os.getcwd(),proc_dir, word_save))
    doc.Range(0,0).InsertAfter(approval_letter)
    word.ActiveDocument.SaveAs2(word_save)
    word.Visible = 1
    print('Please review and modify the Approval Letter. SAVE ANY MODIFICATIONS, and print the letter. DO NOT EXIT WORD.')
    continue_on()


if __name__=='__main__':
    proc_dir = "processing"
    date_save = datetime.strftime(datetime.today(), "%m.%d.%Y")
    get_client_dict()
