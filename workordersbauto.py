import csv
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import collections
import time
import re
import cProfile
#translate pdf123458
def translatepdf():
    """
    use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob('/Users/shaneshimizu/Downloads/*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    validatefile = input("name this file: ")
    schoolcode = input("School Code: ")
    print("\n")
    #validatefile = input("Is " + latest_file + " the file you want to process?" + "\n type yes or no: ")

    if validatefile.lower() != None:
        dateRange = validatefile
        #dateRange = input("Please enter the date range for this report (use underscore instead of spaces): ")
        try:
            data = read_pdf(latest_file, pages = 'all')
            tabula.convert_into(latest_file, f'Work_Order_{dateRange}_EXP.csv',  guess=False, stream=True, area = (18.05,17.9,568.49,756.57), output_format="csv", pages = 'all')
        except:
            print("not a fleet report, check recent download")
            return
        workbookname = f'Work_Order_{dateRange}_EXP.csv'
        time.sleep(1)
        executeAutomation(workbookname, schoolcode)
    else:
        exit()
workbookname = 'Work_Order_qwe_EXP.csv'
schoolcode = 'SW'
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
            rowcount = 0
            for column in readfile:
                hastotal = "Department Totals"
                accountnum = column[0]
                diclist = []
                try:
                    accountnum2 = column[5]
                except:
                    continue
                notaccount = "YAMAHA"
                departmentname = "Department:"
                #if accountnum and department name different, duplicate account with special charcter, append to dictionary
                if accountnum[13:20] in dicEachAccountTotals:
                    theaccount2 = accountnum[13:20]
                    dicEachAccountTotals[str(theaccount2) + 'd'] = []
    
                elif schoolcode.upper() in accountnum and notaccount not in accountnum:
                    theaccount = accountnum[13:20]
                    dicEachAccountTotals[theaccount] = None
                    print(dicEachAccountTotals)
                try:
                    if hastotal in accountnum:
                        try:
                            allfloats = re.findall("[+-]?\d+\.\d+", accountnum2)
                            thetotal = float(allfloats[3])
                        except:
                            accountnum10 = column[8]
                            allfloats7 = re.findall("[+-]?\d+\.\d+", accountnum10)
                            thetotal = accountnum10
                        recentkey = list(dicEachAccountTotals.keys())[-1]
                        dicEachAccountTotals[f"{recentkey}"]=thetotal
                        print(dicEachAccountTotals)
                        print("\n")
                except IndexError:
                    print("exception 1 was ran")
                    if hastotal in accountnum:
                        accountnum3 = column[9]
                        allfloats = re.findall("[+-]?\d+\.\d+", accountnum3)
                        dicEachAccountTotals[theaccount] = accountnum3
                    elif hastotal in accountnum:
                        accountnum4 = column[8]
                        allfloats = re.findall("[+-]?\d+\.\d+", accountnum4)
                        dicEachAccountTotals[theaccount] = accountnum4
                    else: 
                        accountnum5 = column[0]
                        allfloats = re.findall("[+-]?\d+\.\d+", accountnum5)
                        thetotal = float(allfloats[-1])
                        dicEachAccountTotals[theaccount] = thetotal
            for k, v in dicEachAccountTotals.items():
                writefile.writerow((schoolcode, k[0:7], '', '5840', '', '', '', 'STATE VEHICLE R/M CHARGE', v))
                writefile2.writerow(("MA", 2302698, '', '0751', '', '', '', 'STATE VEHICLE R/M CHARGE', v))
                    
    
        pprint.pprint(dicEachAccountTotals)

translatepdf()