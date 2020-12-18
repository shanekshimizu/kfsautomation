import csv
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import time
import getpass

username = getpass.getuser()
def translatepdf():
    pass
    """
    use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob(f'/Users/{username}/Downloads/*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    validatefile = input("Name file: ")
    schoolcode = input("School Code: ")
    os.system(f"open '{latest_file}'")

    if validatefile.lower() != None:
        dateRange = validatefile
        try:
            pathToFile = f'/Users/{username}/Desktop/service_billings/tabula_csv/'
            fileName = f'Mail_{dateRange}.csv'
            joinPath = os.path.join(pathToFile, fileName)
            data = read_pdf(latest_file, pages = 'all')
            tabula.convert_into(latest_file, joinPath, output_format="csv", pages = 'all')
        except:
            print("not a fleet report")
            return
        workbookname = joinPath
        time.sleep(1)
        executeAutomation(workbookname, schoolcode, dateRange)
    else:
        exit()


def executeAutomation(workbookname, schoolcode, dateRange):
    #EXP and REV
    with open(workbookname, 'r') as readit:
        readfile = csv.reader(readit)
        next(readfile)
        pathToFile = f'/Users/{username}/Desktop/service_billings/'
        fileName = f'Mail_{dateRange}'
        joinPath = os.path.join(pathToFile, fileName)
        with open(f'{joinPath}_EXP.csv', 'w') as writeit, open(f'{joinPath}_REV.csv', 'w') as writeit2:
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
            
    os.system(f"open '{joinPath}_EXP.csv'")
    os.system(f"open '{joinPath}_REV.csv'")

if __name__ == "__main__":
    translatepdf()