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
    validatefile = input("Name file: ")
    schoolcode = input("School Code: ")

    if validatefile.lower() != None:
        dateRange = validatefile
        try:
            data = read_pdf(latest_file, pages = 'all')
            tabula.convert_into(latest_file, f'Mail_{dateRange}.csv', output_format="csv", pages = 'all')
        except:
            print("not a fleet report")
            return
        workbookname = f'Mail_{dateRange}.csv'
        time.sleep(1)
        executeAutomation(workbookname, schoolcode)
    else:
        exit()


def executeAutomation(workbookname, schoolcode):
    #EXP and REV
    with open(workbookname, 'r') as readit:
        readfile = csv.reader(readit)
        next(readfile)
        with open(f'EXP_{workbookname}', 'w') as writeit, open(f'REV_{workbookname}', 'w') as writeit2:
            writefile = csv.writer(writeit)
            writefile2 = csv.writer(writeit2)
            for column in readfile:
                accountnum = column[2]
                charge = column[5]
                try:
                    if len(accountnum) < 7 and accountnum.isupper() == False:
                        print(accountnum, " - was 6 numbers long, dusplicated first number")
                        firstnum = accountnum[0]
                        finalaccount = ''
                        finalaccount = f"{firstnum}{accountnum}".format(finalaccount)
                    else:
                        finalaccount = accountnum
                    if charge.isupper() == True:
                        raise IndexError
                except IndexError: # the format of the pdf has changed
                    print("document format changed, adjusting...")
                    for column in readfile:
                        for cell in column:
                            if len(cell) == 7:
                                finalaccount = cell
                            if '$' in cell:
                                charge = cell
                writefile.writerow((schoolcode, finalaccount,'', 3700, '', '', '', 'MAIL CHARGE', charge))
                writefile2.writerow(('MA', 2214012, '', 3700, '', '', '', "MAIL CHARGE", charge))   
            
    os.system(f"open 'EXP_{workbookname}'")
    os.system(f"open 'REV_{workbookname}'")

if __name__ == "__main__":
    translatepdf()