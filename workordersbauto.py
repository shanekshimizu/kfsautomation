import csv
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import collections
import time
import re


def translatepdf():
    """
    Use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob('/Users/shaneshimizu/Downloads/*')
    latest_file = max(list_of_files, key=os.path.getctime)
    validateFile = input("name this file: ")
    schoolCode = input("School Code: ")
    
    if validateFile.lower() != None:
        dateRange = validateFile
        #read recent file in download, read only set area, output into a csv file
        try:
            data = read_pdf(latest_file, pages='all')
            tabula.convert_into(latest_file, f'Work_Order_{dateRange}.csv',  guess=False, stream=True, area=(
                18.05, 17.9, 568.49, 756.57), output_format="csv", pages='all')
        #most recent file is not a accepted file for conversion
        except:
            print("not a fleet report, check recent download")
            return
        workBookName = f'Work_Order_{dateRange}.csv'
        time.sleep(1)
        executeAutomation(workBookName, schoolCode)
    else:
        exit()

def executeAutomation(workBookName, schoolCode):
    """
    Read the converted csv, extract account and respective charges only, write into two seperate csv files 
    """
    with open(workBookName, 'r') as readIt:
        readFile = csv.reader(readIt)
        next(readFile)
        with open(f'EXP_{workBookName}', 'w') as writeExpData, open(f'REV_{workBookName}', 'w') as writeRevData:
            writeExp = csv.writer(writeExpData)
            writeRev = csv.writer(writeRevData)
            dicEachAccountTotals = {}

            for row in readFile:
                # should only use cells with these labels
                hasTotal = "Department Totals"
                accountLabel = "Account"

                # Add accounts to dictionary
                accountNum = row[0]
                if accountNum[13:20] in dicEachAccountTotals:
                    theAccount = accountNum[13:20]
                    dicEachAccountTotals[str(theAccount) + 'd'] = []
                elif schoolCode.upper() in accountNum and accountLabel in accountNum:
                    theAccount = accountNum[13:20]
                    dicEachAccountTotals[theAccount] = [None]

                # Add charges to each account in dictionary
                if hasTotal in accountNum:
                    price = row[5]
                    priceSelect = [6, 7, 8, 9]  #other columns prices may be in
                    
                    try:
                        removeComma = price.replace(',', '')
                        allFloats = re.findall("[+-]?\d+\.\d+", removeComma)
                        realPrice = float(allFloats[3]) #grab the price at 3rd index of that row 
                    #price is not in it's usual row
                    except IndexError:
                        for i in priceSelect:
                            try:
                                price = row[i].replace(',', '')
                                if price == 0: #price may sometimes read 0
                                    raise IndexError
                            except IndexError:
                                if price != 0:
                                    removeComma = price.replace(',', '')
                                    allFloats = re.findall("[+-]?\d+\.\d+", removeComma)
                                    realPrice = float(allFloats[-1])
                                    break

                    #append the charge to the most recent account added to dictionary
                    recentkey = list(dicEachAccountTotals.keys())[-1]
                    dicEachAccountTotals[f"{recentkey}"] = float(realPrice)

            #write accounts and charges to REV and EXP files
            for k, v in dicEachAccountTotals.items():
                writeExp.writerow(
                    (schoolCode, k[0:7], '', '5840', '', '', '', 'STATE VEHICLE R/M CHARGE', v))
                writeRev.writerow(
                    ("MA", 2302698, '', '0751', '', '', '', 'STATE VEHICLE R/M CHARGE', v))

        print("\n")
        pprint.pprint(collections.OrderedDict(dicEachAccountTotals))

        #open csv files to double check before submitting
        os.system(f"open 'EXP_{workBookName}'")
        os.system(f"open 'REV_{workBookName}'")


if __name__ == "__main__":
    translatepdf()
