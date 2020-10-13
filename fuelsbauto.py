import csv
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import time
import collections


def translatepdf():
    """
    use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob('/Users/shaneshimizu/Downloads/*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    validateFile = input("name this file: ")
    schoolCode = input("School Code: ")
    print("\n")

    if validateFile.lower() != None:
        dateRange = validateFile
        #read recent file in download, read only set area, output into a csv file
        try:
            data = read_pdf(latest_file, pages = 'all')
            os.system(f"open {latest_file}")
            tabula.convert_into(latest_file, f'Fuel_{dateRange}_EXP.csv', output_format="csv", pages = 'all')
        #most recent file is not a accepted file for conversion
        except:
            print("not a fleet report, check recent download")
            return
        workBookName = f'Fuel_{dateRange}_EXP.csv'
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

            #other rows that account and price may be in
            accountOptions = [3, 2, 1]
            priceOptions = [11, 10, 9, 8]

            #labels reader should ignore
            mail = "MAIL"
            gas = "GAS"
            uhwo = "UHWO"
            yamaha = "YAMAHA"

            for row in readFile:
                for i in accountOptions:
                    #append REGULAR ACCOUNT ONLY
                    if schoolCode in row[i] and mail not in row[i] and gas not in row[1] and uhwo not in row[i] and yamaha not in row[i]:
                        if row[i] in dicEachAccountTotals:
                            for p in priceOptions:
                                try:
                                    dicEachAccountTotals[row[i]].append(float(row[p]))
                                    break
                                except IndexError:
                                    continue
                        else:
                            dicEachAccountTotals[row[i]] = []
                            for p in priceOptions:
                                try:
                                    dicEachAccountTotals[row[i]].append(float(row[p]))
                                    break
                                except IndexError:
                                    continue
                    #append GAS ONLY
                    if gas in row[1] and schoolCode in row[i] and mail not in row[i]:
                        if row[i] + "G" in dicEachAccountTotals:
                            for p in priceOptions:
                                try:
                                    dicEachAccountTotals[row[i] + "G"].append(float(row[p]))
                                    break
                                except IndexError:
                                    continue
                        else:
                            dicEachAccountTotals[row[i] + "G"] = []
                            for p in priceOptions:
                                try:
                                    dicEachAccountTotals[row[i] + "G"].append(float(row[p]))
                                    break
                                except IndexError:
                                    continue

            #some values of all keys
            for k, v in dicEachAccountTotals.items():
                dicEachAccountTotals[k] = sum(v)

            ##write accounts and charges to REV and EXP files
            for k,v in dicEachAccountTotals.items():
                if k[-1] == "G": #gas accounts
                    writeExp.writerow((schoolCode, k.partition(schoolCode)[2][0:8].strip(), '', '3035', '', '', '', 'GAS CHARGE', v))
                else:
                    writeExp.writerow((schoolCode, k.partition(schoolCode)[2][0:8].strip(), '', '3025', '', '', '', 'GAS CHARGE', v))
                writeRev.writerow(("MA", 2221362, '', '0793', '', '', '', 'GAS CHARGE', v))
                
            pprint.pprint(dicEachAccountTotals)

    #open csv files to double check before submitting
    os.system(f"open 'EXP_{workBookName}'")
    os.system(f"open 'REV_{workBookName}'")


if __name__ == "__main__":
    translatepdf()
