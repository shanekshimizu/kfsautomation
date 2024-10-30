import csv
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import collections
import time
import re
import getpass

username = getpass.getuser()
def translatepdf():
    """
    Use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob('/Users/shaneshimizu/Downloads/*')
    latest_file = max(list_of_files, key=os.path.getctime)
    validateFile = input("name this file: ")
    schoolCode = input("School Code: ")
    os.system(f"open '{latest_file}'")
    
    if validateFile.lower() != None:
        dateRange = validateFile
        pathToFile = f'/Users/{username}/Desktop/service_billings/tabula_csv/'
        fileName = f'Work_Order_{dateRange}.csv'
        joinPath = os.path.join(pathToFile, fileName)
        #read recent file in download, read only set area, output into a csv file
        try:
            data = read_pdf(latest_file, pages='all')
            tabula.convert_into(latest_file, joinPath,  guess=False, stream=True, area=(
                18.05, 17.9, 568.49, 756.57), output_format="csv", pages='all')
        #most recent file is not a accepted file for conversion
        except:
            print("not a fleet report, check recent download")
            return
        workBookName = joinPath
        time.sleep(1)
        executeAutomation(workBookName, schoolCode, dateRange)
    else:
        exit()

def executeAutomation(workBookName, schoolCode, dateRange):
    """
    Read the converted csv, extract account and respective charges only, write into two seperate csv files 
    """
    with open(workBookName, 'r') as readIt:
        readFile = csv.reader(readIt)
        next(readFile)
        pathToFile = f'/Users/{username}/Desktop/service_billings/'
        fileName = f'Work_Order_{dateRange}'
        joinPath = os.path.join(pathToFile, fileName)
        with open(f'{joinPath}_123.csv', 'w') as writeExpData, open(f'{joinPath}_456.csv', 'w') as writeRevData:
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
                    try:
                        price = row[4]
                    except IndexError:
                        continue
                    priceOptions = [5, 6, 7, 8, 9]  #other columns prices may be in
                    
                    try:
                        removeComma = price.replace(',', '')
                        allFloats = re.findall("[+-]?\d+\.\d+", removeComma)
                        realPrice = float(allFloats[3]) #grab the price at 3rd index of that row 
                    #price is not in it's usual row
                    except IndexError:
                        for i in priceOptions:
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
                    (schoolCode, k[0:7], '', '####', '', '', '', 'CHARGE', v))
                writeRev.writerow(
                    ("TEST", "######", '', '####', '', '', '', 'CHARGE', v))

        print("\n")
        pprint.pprint(collections.OrderedDict(dicEachAccountTotals))

        #open csv files to double check before submitting
        os.system(f"open '{joinPath}_123.csv'")
        os.system(f"open '{joinPath}_456.csv'")


if __name__ == "__main__":
    translatepdf()
