import csv
import openpyxl
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import time
from pyexcel.cookbook import merge_all_to_a_book
import pandas as pd
import pyautogui
import collections
from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC


def translatepdf():
    """
    use tabula to convert pdf file into .csv
    """
    list_of_files = glob.glob('/Users/username/Downloads/*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    validatefile = input("name this file: ")
    print("\n")
    #validatefile = input("Is " + latest_file + " the file you want to process?" + "\n type yes or no: ")

    if validatefile.lower() != None:
        dateRange = validatefile
        #dateRange = input("Please enter the date range for this report (use underscore instead of spaces): ")
        try:
            data = read_pdf(latest_file, pages = 'all')
            tabula.convert_into(latest_file, f'Test_Invoice_{dateRange}.csv',  guess=False, stream=True, area = (54.85,15.76,775.07,595.18), output_format="csv", pages = 'all')
        except:
            print("not a fleet report, check recent download")
            return
        workbookname = f'Test_Invoice_{dateRange}.csv'
        #time.sleep(1)
        executeAutomation(workbookname)
    else:
        exit()

def executeAutomation(workbookname):
    with open(workbookname, 'r') as readit:
        readfile = csv.reader(readit)
        dicEachAccount = {'0': [], '1': [], '2': [], '6': [], '9': []}
        somelist = [0,1,2,6,9]
        code1 = "CODE1"
        code2 = "CODE2"
        for column in readfile:
            for i in somelist:
                if column[i] and len(column[5]) == 1:
                    dicEachAccount[f'{i}'].append(column[i])
        
        account = dicEachAccount["2"]
        if uhf in account or rcuh in account:
            #seperate the account 
            pass

         
    print("\n")
    pprint.pprint(collections.OrderedDict(dicEachAccount))

if __name__ == "__main__":
    translatepdf()

