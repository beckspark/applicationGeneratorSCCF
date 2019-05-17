# This is the 5th version of the application generator for the Senior Charity Care Foundation. By Stephen Beck, May 2019.


# For this script to open, you'll need Windows 10 and the following programs installed:
# Word
# SumatraPDF

import glob
import pandas as pd
from difflib import get_close_matches
import string
import os
import shutil
import win32com.client as win32
import pywinauto
from pywinauto import *
import subprocess
import datetime
import pyperclip
from time import sleep
from string import Template


# First thing, we've got to set up the directories to know where everything is at. We've got to figure out where the Dropbox directory is, whether through the server directly or through the Dropbox app on Windows.
try:
    dbDir = os.path.join("\\\\10.1.10.201\\Dropbox\\Services")
    if not os.path.exists(dbDir):
        dbDir = (os.path.join("%s","Dropbox" % os.path.expanduser('~')))
        if not os.path.exists(dbDir):
            ok = 'no'
            while ok != "e":
                ok = input('The Dropbox folder doesn\'t seem to be properly set up! Script will now exit. Enter \'e\': ')
            exit


    penDir = os.path.join(dbDir,"Pending_Applications")
    procDir = os.path.join(penDir,"PROCESSING")
    sumatra = 'C:\\Program Files\\SumatraPDF\\SumatraPDF.exe'
    date = datetime.datetime.today().strftime('%m/%d/%Y')
    dateSave = datetime.datetime.today().strftime('%m.%d.%Y')
    yearSave = datetime.datetime.today().strftime('%Y')
    letterhead = os.path.join(penDir,"Application_Template.docx")

    # Next, it's time to start identifying the file to work on, before we get into the first function
    if not glob.glob(os.path.join(penDir,"*.pdf")):
        cont = 'no'
        while cont != 'y':
            cont = input('Cannot find a .pdf file to process. enter \'y\' to exit: ')
            exit
    else:
        for file in os.listdir(penDir):
            if file.endswith(".pdf"):
                appFile = os.path.join(file)
                appAbs = os.path.join(penDir,appFile)
    processApp = 'no'
    while processApp != 'YES':
        processApp = input('Welcome to the appGen script for SCCF! The pdf file found is ' + appFile + '. Type \"YES\" to continue with this application: ')
    shutil.copy(appAbs,os.path.join(procDir,appFile))
    os.remove(appAbs)
    appAbs = os.path.join(procDir,appFile)

    # First thing is to open the application file as a subprocess with SumatraPDF.
    pdfOpen = subprocess.Popen([sumatra, appAbs], shell=True)
    print('The application will now open in a new window, please refer to it and DO NOT EXIT for the following questions:')
    lname = input('What is the client\'s last name?: ')
    fname = input('What is the client\'s first name?: ')
    initials = fname[0] + lname[0]
    # This next little nifty part will ask for the name of the client's housing, then check it against a database of housing in Utah. If a match is found, the address, zip code, living type, city and County will all be saved automatically.
    db = pd.read_excel(os.path.join(penDir,'Senior_Housing_2019.xlsx'), sheet_name=0)
    dbhousing = db['Housing_Name'].tolist()
    dbaddress = db['Address 1'].tolist()
    dbcities = db['City'].tolist()
    dbzips = db['Zip'].tolist()
    dbtype = db['Housing_Type'].tolist()
    dbcounties = db['County'].tolist()
    housing = input("What is the name of their housing?: ").translate(str.maketrans('', '', string.punctuation))
    if len(get_close_matches(housing, dbhousing)) > 0:
        matches = get_close_matches(housing, dbhousing)
        for m in range(len(matches)):
            yn = input("Did you mean %s instead? Enter Y if yes, or N if no: " % get_close_matches(housing, dbhousing)[m])
            if yn == "Y":
                match = dbhousing.index(get_close_matches(housing, dbhousing)[m])
                housing = str(dbhousing[match])
                address = str(dbaddress[match])
                city = str(dbcities[match])
                zipCode = str(dbzips[match])
                housingType = str(dbtype[match])
                county = str(dbcounties[match])
                break
    elif (len(get_close_matches(housing, dbhousing)) == 0) or (yn == "N"):
            print('%s was not found in the database. Please enter the information for their housing manually: ' % housing)
            address = str(input('What is their street name and number?: '))
            city = str(input('What city?: '))
            zipCode = str(input('What zip code?: '))
            housingType = 'no'
            while housingType != "AL" and housingType != "IL":
                housingType = input('What is their housing type? AL for Assisted Living, IL for Independent Living, please: ')
            county = input('What is the name of their county?: ').replace("County","").replace("county","")
            db1 = db.append(pd.DataFrame([[housing,address,city,zipCode,housingType,county]], columns=db.columns))
            db1.to_excel('Senior_Housing_2019.xlsx', index=False)
    service = input('What is the service that they need? (\'D\' for dental, \'V\' for vision, \'H\' for hearing, \'DV\' for dental and vision, etc.: ')
    servicel = []
    if 'D' in service:
        servicel.append('Dental')
    if 'H' in service:
        servicel.append('Hearing')
    if 'V' in service:
        servicel.append('Vision')
    appDate = input('What is the date of their application? MM/DD/YYYY format only, please: ')
    hh = '3'
    while hh != '1' and hh != '2':
        hh = input('What is their household size? 1 for single, 2 for married only, please: ')
    apt = str(input('What is their apartment number? Number only, please: '))
    phone = str(input('What is their phone number?: ').replace("(", "").replace(")", "").replace("-", ""))
    dob = input('What is their date of birth? MM/DD/YYYY format only, please: ')
    mincome = float(input('What is their MONTHLY income?: ').replace(",", "").replace("$", ""))
    if housingType == "AL":
        facilityCharges = float(input('What do they pay to the facility every month?: '))
        mincome = float(facilityCharges - mincome)
        if mincome < 0:
            mincome = float('45')
    elif housingType == "IL":
        rent = input('What do they pay MONTHLY for rent?: ').replace(",", "").replace("$", "")
    aincome = mincome*12
    food = input('What do they pay MONTHLY for food?: ').replace(",", "").replace("$", "")
    trans = input('What do they pay MONTHLY for transportation?: ').replace(",", "").replace("$", "")
    meds = input('What do they pay MONTHLY for medications?: ').replace(",", "").replace("$", "")
    yourName = str(input('What is YOUR full name?: '))
    title = str(input('What is your title at SCCF?: '))
    # Now we'll take all the data and see whether to approve or deny the application
    subprocess.call(['taskkill', '/f', '/im', 'SumatraPDF.exe'])
    pdfOpen.kill()
    # But first, let's be sure to rename the application so that it gets moved with the rest of the files.
    shutil.move(appAbs,os.path.join(procDir,(lname + "_" + fname + "_Application_" + dateSave + ".pdf")))
    print('Processing data now, please wait...')
    if hh == '1':
        if mincome <= 1011.67:
            hourly = 55
            denture = 53
            flipper = 19
            valplast = 5
            percent = 80
            approval = 'Approval'
        elif mincome <= 1345.52:
            hourly = 63
            denture = 61
            flipper = 22
            valplast = 6
            percent = 77
            approval = 'Approval'
        elif mincome <= 1517.5:
            hourly = 71
            denture = 69
            flipper = 25
            valplast = 7
            percent = 74
            approval = 'Approval'
        elif mincome <= 2023.34:
            hourly = 82
            denture = 80
            flipper = 29
            valplast = 8
            percent = 70
            approval = 'Approval'
        elif mincome > 2023.34:
            approval = 'Denial'
    elif hh == '2':
        if mincome <= 1371.67:
            hourly = 55
            denture = 53
            flipper = 19
            valplast = 5
            percent = 80
            approval = 'Approval'
        elif mincome <= 1824.34:
            hourly = 63
            denture = 61
            flipper = 22
            valplast = 6
            percent = 77
            approval = 'Approval'
        elif mincome <= 2057.5:
            hourly = 71
            denture = 69
            flipper = 25
            valplast = 7
            percent = 74
            approval = 'Approval'
        elif mincome <= 2743.34:
            hourly = 82
            denture = 80
            flipper = 29
            valplast = 8
            percent = 70
            approval = 'Approval'
        elif mincome > 2743.34:
            approval = 'Denial'

    if 'D' in service:
        dentures = denture*2
        if approval == 'Approval':
            approvalIn = open('dentalApproval.txt')
            src = Template(approvalIn.read())
            d = {'date':date,'fname':fname,'lname':lname,'housing':housing,'address':address,'apt':apt,'city':city,'zipCode':zipCode,'phone':phone,'hourly':hourly,'denture':denture,'dentures':dentures,'flipper':flipper,'valplast':valplast,'percent':percent,'yourName':yourName,'title':title}
            approvalLetter = src.substitute(d)
            wordSave = os.path.join(procDir, lname + '_' + fname + '_Dental_Approval_Letter_' + dateSave + '.docx')
            word = win32.DispatchEx('Word.Application')
            doc = word.Documents.Open(letterhead)
            word.Visible = 0
            rng = doc.Range(0,0)
            rng.InsertAfter(approvalLetter)
            word.ActiveDocument.SaveAs2(str(wordSave))
            word.Visible = 1
            cont = 'no'
            while cont != 'YES':
                cont = input('Please review and modify the Approval Letter. SAVE ANY MODIFICATIONS, and print the letter. DO NOT EXIT WORD. Enter YES when ready: ')
            pdfSave = os.path.join(procDir, lname + '_' + fname + '_Dental_Approval_Letter_' + dateSave + '.pdf')
            word.Visible = 0
            doc.ExportAsFixedFormat(pdfSave, 17)
            word.Quit()

        # Establish if this client has a directory in the current year. If not, the script will create one.
        clientDir = os.path.join(dbDir, "Approval letters for Charity Care", yearSave, (county + ' County Approval Letters'), (lname + "_" + fname + "_" + dateSave))
        if not os.path.exists(clientDir):
            os.makedirs(clientDir)

        # Next, we'll move everything with the client's name in it from the procDir to the clientDir
        clientFiles = (lname + "_" + fname + "*")
        for files in glob.glob(os.path.join(procDir,clientFiles)):
            if not os.path.isdir(files):
                shutil.copy(files,clientDir)
                os.remove(files)

        # Now comes time for the fun DentiMax hacky code!
            cont = 'maybe'
            while cont != 'YES' and cont != 'NO':
                cont = input('Do you want to add the client to Dentimax? Enter \'YES\' or \'NO\': ')
            if cont == 'YES':
                dMaxFile = os.path.join(penDir,'cpub-dentimax-APP7-CmsRdsh.rdp')
                subprocess.call(['taskkill','/f','/im', 'mstsc.exe', '/t'])
                dMax = subprocess.Popen(['mstsc',dMaxFile], shell=True)
                cont = 'no'
                while cont != 'YES':
                    cont = input('DentiMax will now start in a separate window. Please log in and navigate to the home page. Enter \'YES\' to continue: ')
                cont = 'no'
                while cont!= 'YES':
                    cont = input('Please wait now. DO NOT TOUCH THE KEYBOARD. We will now search for the client\'s last name. Enter \'YES\' when ready: ')
                app = application.Application()
                h = pywinauto.findwindows.find_windows(title_re=".*GATEWAY1",class_name="RAIL_WINDOW")
                if len(h) > 1:
                    app.connect(handle=h[1])
                elif len(h) <= 1:
                    app.connect(handle=h[0])
                dmWin = app.top_window()
                dmWin.type_keys('{F4}') # Returns to home screen in DentiMax just in case
                sleep(3)
                dmWin.type_keys('{F11}')
                sleep(3)
                dmWin.type_keys('{TAB}')
                dmWin.type_keys('{TAB}')
                dmWin.type_keys('{TAB}')
                dmWin.type_keys(lname, with_spaces=True)
                dmWin.type_keys('{TAB}')
                # If their last name doesn't show up, start a-processin'
                cont = 'no'
                while cont != 'YES' and cont != 'NO':
                    cont = input('Are they already in DentiMax? \'YES\' or \'NO\': ')
                    if cont == 'NO':
                        dmWin.type_keys('{F4}') # Returns to home screen in DentiMax just in case
                        sleep(3)
                        dmWin.type_keys('{F11}')
                        sleep(3)
                        dmWin.type_keys('{F8}')
                        sleep(3)
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys(lname, with_spaces=True)
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys(fname, with_spaces=True)
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys(housing, with_spaces=True)
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys(apt)
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys(zipCode)
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys(phone)
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        pyperclip.copy(dob)
                        dmWin.type_keys('^v')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('{TAB}')
                        dmWin.type_keys('Application approved and processed - '+yourName+' '+dateSave+'. Hourly rate is $'+str(hourly), with_spaces=True)
                        dmWin.type_keys('{F3}')
                    elif cont == "NO":
                        print('This client is already in DentiMax and will not be added at this time. Please remember to add them to a waiting list or give them an appointment, as appropriate.')


    if 'V' in service:
        if percent != 80:
            percent = 70
        if approval == 'Approval':
            approvalIn = open('visionApproval.txt')
            src = Template(approvalIn.read())
            d = {'date':date,'fname':fname,'lname':lname,'housing':housing,'address':address,'apt':apt,'city':city,'zipCode':zipCode,'phone':phone,'yourName':yourName,'title':title,'percent':percent}
            visionLetter = src.substitute(d)
            wordSave = os.path.join(procDir, lname + '_' + fname + '_Approval_Letter_Vision_' + dateSave + '.docx')
            word = win32.DispatchEx('Word.Application')
            doc = word.Documents.Open(letterhead)
            word.Visible = 0
            rng = doc.Range(0,0)
            rng.InsertAfter(visionLetter)
            word.ActiveDocument.SaveAs2(str(wordSave))
            word.Quit()

        # Establish if this client has a directory in the current year. If not, the script will create one.
        clientDir = os.path.join(dbDir, "Approval letters for Charity Care", yearSave, (county + ' County Approval Letters'), (lname + "_" + fname + "_" + dateSave))
        if not os.path.exists(clientDir):
            os.makedirs(clientDir)

        # Next, we'll move everything with the client's name in it from the procDir to the procVision directory.
        procV = os.path.join(dbDir,"Pending_Vision_Applications")
        clientFiles = (lname + "_" + fname + "*")
        for files in glob.glob(os.path.join(procDir,clientFiles)):
            if not os.path.isdir(files):
                shutil.copy(files,procV)
                os.remove(files)

    if 'H' in service:
        if percent != 80:
            percent = 70
        if approval == 'Approval':
            approvalIn = open('hearingApproval.txt')
            src = Template(approvalIn.read())
            d = {'date':date,'fname':fname,'lname':lname,'housing':housing,'address':address,'apt':apt,'city':city,'zipCode':zipCode,'phone':phone,'yourName':yourName,'title':title,'percent':percent}
            hearingLetter = src.substitute(d)
            wordSave = os.path.join(procDir, lname + '_' + fname + '_Hearing_Approval_Letter_' + dateSave + '.docx')
            word = win32.DispatchEx('Word.Application')
            doc = word.Documents.Open(letterhead)
            word.Visible = 0
            rng = doc.Range(0,0)
            rng.InsertAfter(hearingLetter)
            word.ActiveDocument.SaveAs2(str(wordSave))
            word.Quit()

        # Establish if this client has a directory in the current year. If not, the script will create one.
        clientDir = os.path.join(dbDir, "Approval letters for Charity Care", yearSave, (county + ' County Approval Letters'), (lname + "_" + fname + "_" + dateSave))
        if not os.path.exists(clientDir):
            os.makedirs(clientDir)

        # Next, we'll move everything with the client's name in it from the procDir to the procHearing directory.
        procH = os.path.join(dbDir,"Pending_Hearing_Applications")
        clientFiles = (lname + "_" + fname + "*")
        for files in glob.glob(os.path.join(procDir,clientFiles)):
            if not os.path.isdir(files):
                shutil.copy(files,procH)
                os.remove(files)

    elif approval == 'Denial':
        approvalIn = open('denial.txt')
        src = Template(approvalIn.read())
        d = {'date':date,'fname':fname,'lname':lname,'housing':housing,'address':address,'apt':apt,'city':city,'zipCode':zipCode,'phone':phone,'yourName':yourName,'title':title}
        denialLetter = src.substitute(d)
        wordSave = os.path.join(procDir, lname + '_' + fname + '_Denial_Letter_' + dateSave + '.docx')
        word = win32.DispatchEx('Word.Application')
        doc = word.Documents.Open(letterhead)
        word.Visible = 0
        rng = doc.Range(0,0)
        rng.InsertAfter(denialLetter)
        word.ActiveDocument.SaveAs2(str(wordSave))
        word.Visible = 1
        cont = 'no'
        while cont != 'YES':
            cont = input('Please review and modify the Denial Letter. SAVE ANY MODIFICATIONS, and print the letter. DO NOT EXIT WORD. Enter YES when ready: ')

        # Establish if this client has a directory in the current year. If not, the script will create one.
        clientDir = os.path.join(dbDir, "Approval letters for Charity Care", yearSave, (county + ' County Approval Letters'), (lname + "_" + fname + "_" + dateSave))
        if not os.path.exists(clientDir):
            os.makedirs(clientDir)

        # Next, we'll move everything with the client's name in it from the procDir to the Client Files directory.
        clientFiles = (lname + "_" + fname + "*")
        for files in glob.glob(os.path.join(procDir,clientFiles)):
            if not os.path.isdir(files):
                shutil.copy(files,procH)
                os.remove(files)

    # After all that, it's time to read the Client_Data_Master and append it with the information on the new client
    # If the file doesn't exist for this year, add it
    reportTemplate = "\\\\10.1.10.201\\Dropbox\\Services\\Approval Letters for Charity Care\\Charity_Care_Report_Template.xlsx"
    reportFile = os.path.join("\\\\10.1.10.201\\Dropbox\\Services\\Approval Letters for Charity Care\\", yearSave, "Charity_Care_Data_Master_" + yearSave + ".xlsx")
    if not os.path.isfile(reportFile):
        shutil.copyfile(reportTemplate, reportFile)

    # Next, we'll read the dataframe with pandas and append.
    mdb = pd.read_excel(reportFile, sheet_name=0)
    mdb1 = mdb.append(pd.DataFrame([[date, lname, fname, initials, servicel, housing, city, zipCode, mincome, aincome, rent, meds, food, trans, yourName, hourly]], columns=mdb.columns))
    mdb1.to_excel(reportFile, index=False)
    shutil.copy(reportFile,os.path.join(dbDir,"Approval Letters for Charity Care",yearSave,("Charity_Care_Data_Master_"+yearSave+"_EDITABLE.xlsx")))

    # And finally, the ending message and process to exit!
    exit = 'no'
    while exit != 'e':
        exit = input('This client has been processed! Please enter \'e\' to exit: ')
        exit()
except:
    exit = 'no'
    while exit != 'e':
        exit = input('There appears to have been an error. Please try again. Enter \'e\' to exit: ')
    shutil.copy(appAbs,penDir)
    os.remove(appAbs)
    exit()

