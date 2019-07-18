import pandas as pd
import os
import win32com.client as win32
import shutil
from app_gen import continue_on, exit_now
from client_data import *


def get_dict():
    global df
    df = pd.read_csv(os.path.join('processing','client.csv')).to_dict(orient='records')
    df = df[0]


def d_letter():
    global letter
    approval_template = open('dental_approval.txt')
    src = Template(approval_template.read())
    d = {'date':df['approval_date'],'fname':df['fname'],'lname':df['lname'],'housing':df['housing'],'address':df['address'],'apt':df['unit'],'city':df['city'],'zipCode':df['zip_code'],'phone':df['phone'],'hourly':df['hourly'],'denture':df['denture'],'dentures':df['dentures'],'flipper':df['partial'],'valplast':df['valplast'],'percent':df['percent'],'yourName':yourName,'title':title}
    letter = src.substitute(d)
    v = {'date':df['approval_date'],'fname':df['fname'],'lname':df['lname'],'housing':df['housing'],'address':df['address'],'apt':df['unit'],'city':df['city'],'zipCode':df['zip_code'],'phone':df['phone'], 'percent':df['percent'],'yourName':yourName,'title':title}
    v_letter = src.substitute(v)


def letter_word(s, l, nl, nf, d):
    letterhead = os.path.join('processing','application_template.docx')
    letter_save = nl + "_" + nf + "_" + s + "_Approval_Letter_" + d
    word_save = letter_save + ".docx"
    pdf_save = letter_save + ".pdf"
    shutil.copy(letterhead, os.path.join('processing', word_save))
    word = win32.DispatchEx('Word.Application')
    doc = word.Documents.Open(os.path.join(os.getcwd(),'processing', word_save))
    doc.Range(0,0).InsertAfter(l)
    word.ActiveDocument.SaveAs2(word_save)
    word.Visible = 1
    print('Please review and modify the Approval Letter. SAVE ANY MODIFICATIONS, and print the letter. DO NOT EXIT WORD.')
    continue_on()
    word.Visible = 0
    doc.ExportAsFixedFormat(os.path.join(os.getcwd(),'processing',pdf_save), 17)
    word.Quit()



def dental_letter():
    d_letter()
    letter_word()

if __name__=='__main__':
    proc_dir = "processing"
    date_save = datetime.strftime(datetime.today(), "%m.%d.%Y")
    get_client_dict()
    dental_letter()
