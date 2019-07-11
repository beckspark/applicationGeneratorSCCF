import os
import glob
import pandas as pd
import csv
import string
from datetime import datetime, timedelta
from difflib import get_close_matches
from __init__ import *


def poverty_lookup():
    global poverty_1, poverty_2
    df = pd.read_html(io='https://aspe.hhs.gov/poverty-guidelines',match='POVERTY GUIDELINES FOR THE 48 CONTIGUOUS STATES AND THE DISTRICT OF COLUMBIA')
    df = df[0].drop([0])
    df.columns = df.columns.to_flat_index()
    df = df.rename(index=str, columns={df.columns[0]: "Persons", df.columns[1]: "Guideline"})
    poverty_1 = float((df[df["Persons"] == '1']["Guideline"].item()).replace("$","").replace(",",""))
    poverty_2 = float((df[df["Persons"] == '2']["Guideline"].item()).replace("$","").replace(",",""))


def approve_client():
    global approval, hourly, percent, denture, dentures, partial, valplast
    poverty_lookup()
    if status == '1':
        poverty = poverty_1
    elif status == '2':
        poverty = poverty_2
    limit = poverty*2
    percent = 0
    approval = 'waiting'
    if aincome > limit or age < '55':
        approval = 'Denial'
    elif aincome <= poverty:
        hourly = 55
        percent = 20
        denture = 53
        dentures = 106
        partial = 19
        valplast = 5
    elif aincome <= (poverty*1.33):
        hourly = 63
        denture = 61
        dentures = 122
        partial = 22
        valplast = 6
    elif aincome <= (poverty*1.5):
        hourly = 71
        denture = 69
        dentures = 138
        partial = 25
        valplast = 7
    elif aincome <= (poverty*2):
        hourly = 82
        denture = 80
        dentures = 160
        partial = 29
        valplast = 8
    if percent != 20:
        percent = 30
    if approval != 'Denial':
        approval = 'Approval'


def housing_lookup():
    global housing, address, city, zip_code, housing_type, county
    db = pd.read_csv('senior_housing.csv')
    h = input('What is the name of their housing?: ').translate(str.maketrans("", "", string.punctuation))
    matches = get_close_matches(h,db['name'])
    if len(get_close_matches(h,db['name'])) > 0:
           for m in range(len(matches)):
               dbl = db[db.name == matches[m]]
               housing = matches[m]
               address = dbl.address1.item()
               city = dbl.city.item()
               zip_code = str(dbl.zip_code.item())
               housing_type = dbl.housing_type.item()
               county = dbl.county.item()
               yn = input('Do you mean %s at %s in %s? Enter \'y\' for yes, \'n\' for no: ' % (housing, address, city))
               if yn == 'y':
                   break
    if (len(matches) == 0) or yn == 'n':
        print('This location was not found in the database. Please enter the information for their housing manually: ')
        address = input('What is the address of their housing?: ')
        city = input('What is the name of the city?: ')
        zip_code = input('What is the zip code?: ')
        housing_type = 'unknown'
        while housing_type != 'AL' and housing_type != 'IL':
            housing_type = input('What is the housing type? \'AL\' for Assisted Living, \'IL\' for Independent Living, please: ')
        county = input('What is the name of the county?: ').replace("County","").replace("county","")
        db = db.append(pd.DataFrame([[h,address,city,zip_code,housing_type,county]], columns=db.columns))
        db.to_csv('senior_housing.csv', index = False)


def service_data():
    global service
    service = input('What is the service that they need? (\'D\' for dental, \'V\' for vision, \'H\' for hearing, \'DV\' for dental and vision, etc.: ')
    servicel = ""
    if 'D' in service:
        servicel = servicel + ', Dental'
    if 'H' in service:
        servicel = servicel + ', Hearing'
    if 'V' in service:
        servicel = servicel + ', Vision'
    service = servicel[2:]


def find_age():
    global age
    tday = datetime.today()
    bday = datetime.strptime(dob, "%m/%d/%Y")
    age = str(tday.year - bday.year - ((tday.month, tday.day) < (bday.month, bday.day)))


def data_to_csv(d, t):
    proc = 'processing'
    d.to_csv(os.path.join(os.getcwd(),proc,(t+'.csv')), index = False)


def format_data():
    global sccf_clients_clients, sccf_clients_demographics, approval_date, full_name, initials, aincome
    approval_date = datetime.strftime(datetime.today(),"%m/%d/%Y")
    full_name = lname + ", " + fname
    initials = fname[0] + lname[0]
    aincome = float(mincome * 12)
    find_age()
    approve_client()
    sccf_clients_clients = pd.DataFrame({'apply_date':apply_date, 'approval_date':approval_date, 'lname':lname,'fname':fname, 'full_name':full_name, 'initials':initials, 'service':service, 'hourly':hourly, 'phone':phone, 'housing':housing, 'address':address, 'city':city, 'unit':unit, 'zip_code':zip_code, 'county':county, 'housing_type':housing_type, 'approval':approval, 'percent':percent, 'denture':denture, 'dentures':dentures, 'partial':partial, 'valplast':valplast, 'dob':dob}, index=[0])
    sccf_clients_demographics = pd.DataFrame({'full_name':full_name, 'sex':sex, 'age':age, 'resident':resident, 'zip_code':zip_code, 'edu':edu, 'aincome':aincome, 'dental_insurance':dental_insurance, 'ins_type':ins_type, 'medicaid':medicaid, 'disability':disability, 'dis_type':dis_type, 'race':race, "lang":lang}, index=[0])
    data_to_csv(sccf_clients_clients, ('client'))
    data_to_csv(sccf_clients_demographics, ('demo'))


def housing_type_data():
    global rent, food, transp, util, fac_charges, mincome
    if housing_type == "IL":
        rent = input('What is their monthly charge for rent?: ')
        food = input('What do they pay for food every month?: ')
        transp = input('What do they pay for transportation every month?: ')
        util = input('What do they pay for utilities every month?: ')
        fac_charges = 'n/a'
    elif housing_type == "AL":
        fac_charges = float(input('What do they pay to the facility every month?: ').replace("$","").replace(",",""))
        mincome = mincome - fac_charges
        rent = 'n/a'
        food = 'n/a'
        transp = 'n/a'
        util = 'n/a'


def relationship_status():
    global status
    status = 'nah'
    while status != '1' and status != '2':
        status = input('What is their relationship status?: \'1\' for single/divorced/widowed/etc, \'2\' for married/cohabiting/etc.: ')


def get_demo_data():
    global sex, resident, edu, dental_insurance, ins_type, medicaid, disability, dis_type, race, latinx, lang
    print('\n The next questions are for demographic purposes. If they didn\'t include a demographic form, please ask for this information before you continue...\n')
    continue_on()
    sex = input('What is their sex? \'1\' for Male, \'2\' for Female, \'3\' for anything else: ')
    resident = input('Are they a Utah resident? \'0\' for no, \'1\' for yes: ')
    edu = input('What is their education level? \'1\' for less than High School, \'2\' for High School, \'3\' for Bachelor\'s (or higher), \'4\' for Unknown: ')
    dental_insurance = input('Do they have dental insurance? \'0\' for no, \'1\' for yes: ')
    if dental_insurance == '1':
        ins_type = input('Please list their insurance: ')
    elif dental_insurance == '0':
        ins_type = 'n/a'
    medicaid = input('Do they have Medicaid? \'0\' for no, \'1\' for yes: ')
    disability = input('Do they have a disability? \'0\' for no, \'1\' for yes: ')
    if disability == '1':
        dis_type = input('Which disabilities? (a)mbulatory, (c)ognitive, (h)earing, (i)ndependent living, (s)elf-care, (v)ision, (o)ther, (u)nreported/unknown (for example, enter \'aci\' for someone with ambulatory, cognitive & independent living disabilities): ')
    elif disability == '0':
        dis_type = 'n/a'
    race = input('What race are they? \'1\' American Indian/Alaskan Native, \'2\' Asian, \'3\' Native Hawaiian or Pacific Islander, \'4\' White/Caucasian, \'5\' Two or more races \'6\' Unreported/Unknown: ')
    latinx = input('Are they hispanic/latinx? \'0\' for no, \'1\' for yes: ')
    lang = input('What is their primary language? \'1\' for English, \'2\' for Spanish, \'3\' for French, \'4\' for German, \'5\' for Chinese, \'6\' for Russian, \'7\' for Other, \'8\' for Unknown: ')


def get_data():
    global fname, lname, apply_date, mincome, unit, phone, dob, meds, phys, clothes, other
    fname = input('What is the client\'s first name?: ')
    lname = input('What is the client\'s last name?: ')
    apply_date = input('What is the date of their application? MM/DD/YYYY format only, please: ')
    housing_lookup()
    mincome = float(input('What is their monthly income?: ').replace("$","").replace(",",""))
    relationship_status()
    unit = input('What is their apartment or unit number?: ')
    phone = input('What is their phone number?: ')
    dob = input('What is their date of birth? MM/DD/YYYY format only, please: ')
    service_data()
    housing_type_data()
    meds = input('What do they pay for medications every month?: ')
    phys = input('What do they pay to their physicians every month?: ')
    clothes = input('What do they pay for clothing every month?: ')
    other = input('What do they pay for other expenses every month?: ')
    get_demo_data()
    format_data()


if __name__=='__main__':
    get_data()
