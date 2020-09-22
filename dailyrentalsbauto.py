import csv
import openpyxl
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import time
from pyexcel.cookbook import merge_all_to_a_book
import pandas as pd
import pyautogui
import collections
import time



def translatepdf():
    pass
    """
    use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob('/Users/shaneshimizu/Downloads/*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    validatefile = input("name this file: ")
    schoolcode = input("School Code: ")
    #validatefile = input("Is " + latest_file + " the file you want to process?" + "\n type yes or no: ")

    if validatefile.lower() != None:
        dateRange = validatefile
        #dateRange = input("Please enter the date range for this report (use underscore instead of spaces): ")
        data = read_pdf(latest_file, pages = 'all')
        tabula.convert_into(latest_file, f'Daily_Rental_{dateRange}_EXP.csv', output_format="csv", pages = 'all')
        workbookname = f'Daily_Rental_{dateRange}_EXP.csv'
        time.sleep(1)
        executeAutomation(workbookname, schoolcode)
    else:
        exit()

def executeAutomation(workbookname, schoolcode):
    with open(workbookname, 'r') as readit:
        readfile = csv.reader(readit)
        next(readfile)
        with open(f'EXP_{workbookname}.csv', 'w') as writeit, open(f'REV_{workbookname}.csv', 'w') as writeit2:
            writefile = csv.writer(writeit)
            writefile2 = csv.writer(writeit2)
            dicEachAccountTotals = {}
            prices = dicEachAccountTotals.values()
            total = sum(prices)
            for column in readfile:
                accountnum = column[3]
                accountnum2 = column[0]
                hasparen = "("
                if not accountnum:
                    continue
                if not accountnum2:
                    continue
                if hasparen in accountnum:
                    theaccount = accountnum[5:12]
                    charge = column[14]
                    writefile.writerow(("MA", theaccount, '5705', '', '', '', 'DAILY RENTAL CHARGE', charge))
                    writefile.writerow(("MA", theaccount, '0704', '', '', '', 'DAILY RENTAL CHARGE', charge))

                #Failsafe 1
                if schoolcode.upper() in accountnum2:
                    charge = column[11]
                    codehere = accountnum2.find(schoolcode.upper())
                    theaccount = accountnum2[int(codehere) + 3: int(codehere) + 10]
                    if theaccount in dicEachAccountTotals:
                        dicEachAccountTotals[theaccount].append(float(charge))
                    else:
                        dicEachAccountTotals[theaccount] = [float(charge)]
            for k,v in dicEachAccountTotals.items():
                dicEachAccountTotals[k] = sum(dicEachAccountTotals[k])
            for k,v in dicEachAccountTotals.items():
                writefile.writerow((schoolcode, k, '', '5705', '', '', '', 'DAILY RENTAL', v))
                writefile2.writerow((schoolcode, 2302699, '', '0704', '', '', '', 'DAILY RENTAL', v))
    
    sendToKFS()

def sendToKFS():
    pass


                



translatepdf()