import csv
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import time
import collections
import getpass

username = getpass.getuser()
def translatepdf():
    """
    use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob(f'/Users/{username}/Downloads/*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    validateFile = input("name this file: ")
    schoolCode = input("School Code: ")
    os.system(f"open '{latest_file}'")

    if validateFile.lower() != None:
        dateRange = validateFile
        #read recent file in download, read only set area, output into a csv file
        try:
            pathToFile = f'/Users/{username}/Desktop/tabula_csv/'
            fileName = f'File_Name_{dateRange}.csv'
            joinPath = os.path.join(pathToFile, fileName)
            data = read_pdf(latest_file, pages = 'all')
            tabula.convert_into(latest_file, joinPath, guess=False, stream=True, area = (18.05,17.9,568.49,756.57), output_format="csv", pages = 'all')
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
    Read the converted csv, extract account and respective prices only, write into two seperate csv files 
    """
    with open(workBookName, 'r') as readIt:
        readFile = csv.reader(readIt)
        next(readFile)
        pathToFile = f'/Users/{username}/Desktop/service_billings/'
        fileName = f'File_Name_{dateRange}'
        joinPath = os.path.join(pathToFile, fileName)

        with open(f'{joinPath}_123.csv', 'w') as writeExpData, open(f'{joinPath}_456.csv', 'w') as writeRevData:
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
                    if realAccount.strip() in dicEachAccountTotals:
                        dicEachAccountTotals[realAccount.strip()].append(float(price))
                        print("account exists")
                    else:
                        dicEachAccountTotals[realAccount.strip()] = [float(price)]

            #sum values of all keys
            for k,v in dicEachAccountTotals.items():
                dicEachAccountTotals[k] = sum(dicEachAccountTotals[k])

            #write accounts and charges to REV and EXP files
            for k,v in dicEachAccountTotals.items():
                writeEXP.writerow((schoolCode, k, '', '####', '', '', '', 'CHARGE', v))
                writeRev.writerow(('Test', #######, '', '####', '', '', '', 'CHARGE', v))

        pprint.pprint(dicEachAccountTotals)

    #open csv files to double check before submitting
    os.system(f"open '{joinPath}_123.csv'")
    os.system(f"open '{joinPath}_123.csv'")
    

if __name__ == "__main__":
    translatepdf()
