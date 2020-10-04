import csv
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import time


def translatepdf():
    pass
    """
    use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob('/Users/shaneshimizu/Downloads/*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    validatefile = input("name this file ")
    #validatefile = input("Is " + latest_file + " the file you want to process?" + "\n type yes or no: ")

    if validatefile.lower() != None:
        dateRange = validatefile
        #dateRange = input("Please enter the date range for this report (use underscore instead of spaces): ")
        data = read_pdf(latest_file, pages = 'all')
        tabula.convert_into(latest_file, f'Mail_{dateRange}_EXP.csv', output_format="csv", pages = 'all')
        workbookname = f'Mail_{dateRange}_EXP.csv'
        time.sleep(3)
        executeAutomation(workbookname)
    else:
        exit()


def executeAutomation(workbookname):
    #EXP and REV
    with open(workbookname, 'r') as readit:
        readfile = csv.reader(readit)
        next(readfile)
        with open(f'EXP_{workbookname}.csv', 'w') as writeit, open(f'REV_{workbookname}.csv', 'w') as writeit2:
            writefile = csv.writer(writeit)
            writefile2 = csv.writer(writeit2)
            for column in readfile:
                accountnum = column[2]
                charge = column[5]
                if len(accountnum) < 7:
                    firstnum = accountnum[0]
                    finalaccount = ''
                    finalaccount = f"{firstnum}{accountnum}".format(finalaccount)
                writefile.writerow(("MA", finalaccount,'', 3700, '', '', '', 'MAIL CHARGE', charge))
                writefile2.writerow(("MA", 2214012, '', 3700, '', '', '', "MAIL CHARGE", charge))
         

if __name__ == "__main__":
    translatepdf()