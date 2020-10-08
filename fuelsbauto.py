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
    validatefile = input("name this file: ")
    schoolcode = input("School Code: ")
    print("\n")

    if validatefile.lower() != None:
        dateRange = validatefile
        try:
            data = read_pdf(latest_file, pages = 'all')
            tabula.convert_into(latest_file, f'Fuel_{dateRange}_EXP.csv', output_format="csv", pages = 'all')
        except:
            print("not a fleet report, check recent download")
            return
        workbookname = f'Fuel_{dateRange}_EXP.csv'
        time.sleep(1)
        executeAutomation(workbookname, schoolcode)
    else:
        exit()

def executeAutomation(workbookname, schoolcode):

    with open(workbookname, 'r') as readit:
        readfile = csv.reader(readit)
        next(readfile)
        with open(f'EXP_{workbookname}', 'w') as writeit, open(f'REV_{workbookname}', 'w') as writeit2:
            writefile = csv.writer(writeit)
            writefile2 = csv.writer(writeit2)
            dicEachAccountTotals = {}
            accountoptions = [3, 2, 1]
            priceoptions = [11, 10, 9, 8]
            mail = "MAIL"
            gas = "GAS"
            uhwo = "UHWO"
            yamaha = "YAMAHA"
            for column in readfile:
                for i in accountoptions:
                    #REGULAR ACCOUNT ONLY
                    if schoolcode in column[i] and mail not in column[i] and gas not in column[1] and uhwo not in column[i] and yamaha not in column[i]:
                        if column[i] in dicEachAccountTotals:
                            for p in priceoptions:
                                try:
                                    dicEachAccountTotals[column[i]].append(float(column[p]))
                                    break
                                except IndexError:
                                    continue
                        else:
                            dicEachAccountTotals[column[i]] = []
                            for p in priceoptions:
                                try:
                                    dicEachAccountTotals[column[i]].append(float(column[p]))
                                    break
                                except IndexError:
                                    continue
                    #GAS ONLY
                    if gas in column[1] and schoolcode in column[i] and mail not in column[i]:
                        if column[i] + "G" in dicEachAccountTotals:
                            for p in priceoptions:
                                try:
                                    dicEachAccountTotals[column[i] + "G"].append(float(column[p]))
                                    break
                                except IndexError:
                                    continue
                        else:
                            dicEachAccountTotals[column[i] + "G"] = []
                            for p in priceoptions:
                                try:
                                    dicEachAccountTotals[column[i] + "G"].append(float(column[p]))
                                    break
                                except IndexError:
                                    continue

            
            for k, v in dicEachAccountTotals.items():
                dicEachAccountTotals[k] = sum(v)
            for k,v in dicEachAccountTotals.items():
                if k[-1] == "G":
                    writefile.writerow((schoolcode, k.partition(schoolcode)[2][0:8].strip(), '', '3035', '', '', '', 'GAS CHARGE', v))
                else:
                    writefile.writerow((schoolcode, k.partition(schoolcode)[2][0:8].strip(), '', '3025', '', '', '', 'GAS CHARGE', v))
                writefile2.writerow(("MA", 2221362, '', '0793', '', '', '', 'GAS CHARGE', v))
                
           
            pprint.pprint(dicEachAccountTotals)
    os.system(f"open 'EXP_{workbookname}'")
    os.system(f"open 'REV_{workbookname}'")


    
    
if __name__ == "__main__":
    translatepdf()
