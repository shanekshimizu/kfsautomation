import csv
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import time
import collections

def translatepdf():
    pass
    """
    use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob('/Users/shaneshimizu/Downloads/*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    validateFile = input("name this file: ")
    schoolCode = input("School Code: ")

    if validateFile.lower() != None:
        dateRange = validateFile
        #read recent file in download, read only set area, output into a csv file
        try:
            data = read_pdf(latest_file, pages = 'all')
            tabula.convert_into(latest_file, f'Daily_Rental_{dateRange}_EXP.csv', guess=False, stream=True, area = (18.05,17.9,568.49,756.57), output_format="csv", pages = 'all')
        #most recent file is not a accepted file for conversion
        except:
            print("not a fleet report, check recent download")
            return
        workBookName = f'Daily_Rental_{dateRange}_EXP.csv'
        time.sleep(1)
        executeAutomation(workBookName, schoolCode)
    else:
        exit()

def executeAutomation(workBookName, schoolCode):
    """
    Read the converted csv, extract account and respective prices only, write into two seperate csv files 
    """
    with open(workBookName, 'r') as readIt:
        readFile = csv.reader(readIt)
        next(readFile)
        with open(f'EXP_{workBookName}', 'w') as writeExpData, open(f'REV_{workBookName}', 'w') as writeRevData:
            writeEXP = csv.writer(writeExpData)
            writeRev = csv.writer(writeRevData)
            dicEachAccountTotals = {}
            
            for row in readFile:
                accountNum = row[0]
                hasParen = "(" + schoolCode #prevent other words from reading as schoolcode
                if not accountNum:
                    continue
                if hasParen in accountNum:
                    realSchoolCode = accountNum.find(hasParen.upper())
                    realAccount = accountNum[int(realSchoolCode) + 3: int(realSchoolCode) + 11]
                    price = row[10].replace(',', '')
                    if realAccount in dicEachAccountTotals:
                        dicEachAccountTotals[realAccount.strip()].append(float(price))
                    else:
                        dicEachAccountTotals[realAccount.strip()] = [float(price)]

            #sum values of all keys
            for k,v in dicEachAccountTotals.items():
                dicEachAccountTotals[k] = sum(dicEachAccountTotals[k])

            #write accounts and charges to REV and EXP files
            for k,v in dicEachAccountTotals.items():
                writeEXP.writerow((schoolCode, k, '', '5705', '', '', '', 'DAILY RENTAL CHARGE', v))
                writeRev.writerow((schoolCode, 2302699, '', '0704', '', '', '', 'DAILY RENTAL CHARGE', v))

        pprint.pprint(dicEachAccountTotals)

    #open csv files to double check before submitting
    os.system(f"open 'EXP_{workBookName}'")
    os.system(f"open 'REV_{workBookName}'")
    

if __name__ == "__main__":
    translatepdf()