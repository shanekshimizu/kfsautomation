import openpyxl
import csv
import openpyxl
import pprint
import glob
import os
from tabula import read_pdf
import tabula
import time
from pyexcel.cookbook import merge_all_to_a_book


fname = 'Fuel_whatbout_EXP.xlsx'
wb = openpyxl.load_workbook(fname)
sheet = wb["Fuel_whatbout_EXP"]

def testfunction():
#GRAB ALL THE CELLS WITH MA (not gas)
    listofcoordinate = []
    listofaccounts = []
    testaccountlist = []
    for rowOfCellObjects in sheet['##':'####']:
        for cellObj in rowOfCellObjects:
                regulartotal = "##"
                whyaccounthere = None
                strCellObj = str(cellObj.value)
                whatsthis = cellObj.coordinate[1:4]
                bcoordinate = sheet[f"B{whatsthis}"]
                ccoordinate = sheet[f"C{whatsthis}"]
                ccoordinateFinal = str(ccoordinate.value)
                bcoordinateFinal = str(bcoordinate.value)
                regulargastotal = "##"
                if strCellObj == "None" and regulargastotal in bcoordinateFinal:
                    MACoordinate2 = str(cellObj.coordinate)
                    MACoordinateFinal2 = MACoordinate2[1:4]
                    bcellvalue = sheet[f"B{MACoordinateFinal2}"]
                    bcelvaluefinal = bcellvalue.value[-10:]
                    sheet[f"D{MACoordinateFinal2}"] = bcelvaluefinal
                    updatedDcoordinate = sheet[f"D{MACoordinateFinal2}"]
                    listofcoordinate.append(updatedDcoordinate.coordinate)
                    listofaccounts.append(updatedDcoordinate.coordinate)
                    wb.save(fname)
                if strCellObj == "0" and regulargastotal in ccoordinateFinal:
                    MACoordinateC = str(cellObj.coordinate)
                    MACoordinateFinalC = MACoordinateC[1:4]
                    ccellvalue = sheet[f"C{MACoordinateFinalC}"]
                    ccellvaluefinal = ccellvalue.value
                    sheet[f"D{MACoordinateFinalC}"] = ccellvaluefinal
                    updatedCcoordinate = sheet[f"D{MACoordinateFinalC}"]
                    listofcoordinate.append(updatedCcoordinate.coordinate)
                    listofaccounts.append(updatedCcoordinate.coordinate)
                    wb.save(fname)
                if regulartotal in strCellObj:
                    MACoordinate = cellObj.coordinate
                    MAValue = cellObj.value
                    listofcoordinate.append(MACoordinate)
                    listofaccounts.append(MACoordinate)

    #GRAB ALL THE CELLS WITH MA (only gas) and replace B w/ D
    listofGascoordinates = []
    listofgasaccounts = []
    for rowOfCellObjects in sheet['##':'##']:
        for cellObj in rowOfCellObjects:
            regulartotal = "GAS"
            strCellObj = str(cellObj.value)
            if regulartotal in strCellObj:
                MACoordinate = str(cellObj.coordinate)
                MACoordinateFinal = MACoordinate[1:4]
                MAValue = cellObj.value
                MAAccount = cellObj.value
                listofGascoordinates.append(f"D{MACoordinateFinal}")
                listofaccounts.remove(f"D{MACoordinateFinal}")

    #get account D if B has gas    
    for accountCells in sheet[f'##': '####']:
        for cellobjfinal in accountCells:
            strcellobjfinal = cellobjfinal.coordinate
            strcellobjfinalval = cellobjfinal.value
            y = 0
            while y < len(listofGascoordinates):
                #its adding D80-89 because D8 is in D80
                if listofGascoordinates[y] == strcellobjfinal:
                    listofgasaccounts.append(f"{strcellobjfinalval}G")
                y += 1
    
    #get prices for the Gas only
    h = 0
    listofgasprices = []
    listOfFuelPrices = []
    while h < len(listofGascoordinates):
        getValOnly2 = listofGascoordinates[h]
        finalVal2 = getValOnly2[1:4]
        thecells8 = sheet[f'K{finalVal2}']
        if thecells8.value == None:
            #set K cell equal to J cell
            sheet[f"K{finalVal2}"].value = sheet[f"J{finalVal2}"].value
            listofgasprices.append(sheet[f"K{finalVal2}"].value)
            wb.save(fname)
        else:
            fuelpriceCoor8 = thecells8.coordinate
            fuelpriceVal8 = thecells8.value
            listofgasprices.append(fuelpriceVal8)
        h += 1
    #REMOVES ALL account COORDINATES that have the gas
    matchingvalues = []
    c = 0
    while c < len(listofGascoordinates):
        if any(thiscoor in listofcoordinate for thiscoor in listofGascoordinates):
            listofcoordinate.remove(str(listofGascoordinates[c]))
        c += 1


    #GET THE FUEL PRICES after removing gas coordinates
    i = 0
    #listOfFuelPrices = []
    while i < len(listofcoordinate):
        getValOnly = listofcoordinate[i]
        finalVal = getValOnly[1:4]
        thecells = sheet[f'K{finalVal}']
        fuelpriceCoor = thecells.coordinate
        fuelpriceVal = thecells.value
        listOfFuelPrices.append(fuelpriceVal)
        i += 1



    #GET ACCOUNT NUMBER AFTER REMOVING gas
    r = 0
    actualaccountlist = []
    while r < len(listofcoordinate):
        getValOnly2 = listofcoordinate[r]
        finalVal2 = getValOnly2[1:4]
        thecells2 = sheet[f'D{finalVal2}']
        accountcoor2 = thecells2.coordinate
        accountVal2 = str(thecells2.value)
        actualaccountlist.append(accountVal2)
        r += 1


    #FIND EACH UNIQUE INSTANCE OF REGULAR FUEL and put into list
    dicEachAccountTotals = {}
    uniqueAccounts = set(actualaccountlist)
    uniqueAccountsList = list(uniqueAccounts)
    p = 0
    while p < len(uniqueAccountsList):
        dicEachAccountTotals[f"{uniqueAccountsList[p]}"] = 0
        p += 1

    #FIND EACH UNIQUE INSTANCE OF GAS and put into list
    dicEachGasAccountTotals = {}
    uniqueGasAccounts = set(listofgasaccounts)
    uniqueAccountsGasList = list(uniqueGasAccounts)
    s = 0
    while s < len(uniqueAccountsGasList):
        dicEachGasAccountTotals[f"{uniqueAccountsGasList[s]}"] = 0
        s += 1


    #UPDATE THE values IN EACH ACCOUNT after adding all the prices for each account (fuel only)
    q = 0
    while q < len(actualaccountlist):
        if actualaccountlist[q] in dicEachAccountTotals:
            dicEachAccountTotals[f"{actualaccountlist[q]}"] = []
        q += 1

    #UPDATE THE values IN EACH ACCOUNT after adding all the GAS PRICES for each account (gas only)
    u = 0
    while u < len(listofgasaccounts):
        if listofgasaccounts[u] in dicEachGasAccountTotals:
            dicEachGasAccountTotals[f"{listofgasaccounts[u]}"] = []
        u += 1

    #MERGE ACCOUNT AND FUEL LIST
    combinedFuelList = list(zip(actualaccountlist, listOfFuelPrices))
    combinedGasList = list(zip(listofgasaccounts, listofgasprices))

    #PRINT THE 1ST/2ND INDEX OF TUPLE IN LIST
    firstTuple = [u[0] for u in combinedFuelList]
    secondTuple = [u[1] for u in combinedFuelList]

    firstTupleGas = [w[0] for w in combinedGasList]
    secondTupleGas = [w[1] for w in combinedGasList]

    
    #APPEND RESPECTIVE CHARGE TO EACH ACCOUNT (fuel only)
    y = 0
    while y < len(combinedFuelList):
        if firstTuple[y] in dicEachAccountTotals.keys():
            dicEachAccountTotals[firstTuple[y]].append(secondTuple[y])
        y += 1

    #APPEND RESPECTIVE CHARGE TO EACH ACCOUNT (gas only)
    w = 0
    while w < len(combinedGasList):
        if firstTupleGas[w] in dicEachGasAccountTotals.keys():
            dicEachGasAccountTotals[firstTupleGas[w]].append(secondTupleGas[w])
        w += 1

    #GET THE SUM OF EACH ACCOUNT (fuel only)
    for key2 in dicEachAccountTotals:
        dicEachAccountTotals[key2] = sum(dicEachAccountTotals[key2])

    for key3 in dicEachGasAccountTotals:
        dicEachGasAccountTotals[key3] = sum(dicEachGasAccountTotals[key3])

    dicEachAccountTotals.update(dicEachGasAccountTotals)

    #print("\n")
    list_of_files5 = glob.glob('/Users/username/Desktop/VSCode/automationtesting/*') # * means all if need specific format then *.csv
    latest_file5 = sorted(list_of_files5, key=os.path.getctime)
    print("File saved as:", latest_file5[-3])

    with open(latest_file5[-3], "w") as f:
        writer = csv.writer(f, dialect="excel")
        for key, value in dicEachAccountTotals.items():
            writer.writerow(["{} {} {} {} {} {} {} {} {}".format("TEST","######","","####","","","","CHARGE","test")])
    #pprint.pprint(dicEachAccountTotals)
    #print("\n")

testfunction()