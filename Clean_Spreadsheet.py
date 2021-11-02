import xlrd
import xlsxwriter
import math

def correlateValues(posList, freqList): # Correlates each PoS tag to its corresponding frequency value, and creates a dictionary of them
    ldict = {}
    freqListCounter = 0
    if not isinstance(freqList, float or int):
        for j in posList:
            ldict[j] = freqList[freqListCounter]
            freqListCounter += 1
    else:
        for j in posList:
            ldict[j] = freqList
            freqListCounter += 1
    return ldict

def retrData(row, column): # Retrieves data from a given cell, in list form if its not already a number
    value = worksheet.cell(row, column)
    valueCleaned = value.value
    if isinstance(valueCleaned, float or int):
        return [float(valueCleaned)]
    else:
        lValueList = valueCleaned.split(".")
        return lValueList

def delName(dict): # Deletes all entries in given dictionary with key "Name"
    if "Name" in dict:
        del dict["Name"]
        return dict
    else:
        return dict

def writeData(dict, word, iteration, totalFrq): # This function calculates frequency per million words and Zipf Value, then writes the data in correct cells
    keys = list(dict.keys())
    keyCounter = 0
    while dict != {}:
        givenPOS = keys[keyCounter]
        currentRow = iteration + keyCounter
        writerCounter = 0
        # This sets up the needed counters and variables for the rest of the function
        fpmw = float(finalDict[givenPOS]) * (10 ** 6) / 51000000
        zipfV = math.log(fpmw * 1000, 10)
        # This applies the math to the data
        for writer in [word[0], float(totalFrq), givenPOS, float(finalDict[givenPOS]), round((fpmw), 2), round((zipfV), 6)]:
            sheet.write(currentRow, writerCounter, writer)
            writerCounter += 1
        del finalDict[givenPOS]
        keyCounter += 1
    return iteration + keyCounter

readBook = input("Path to read file: ") # Gets the path to the read file
writeBook = input("Path to write file: ") # Gets the path to the write file
sheetname = input("Name of sheet: ")
rowNumber = int(input("Number of rows: "))
headerCounter = 0 # This is used to count iterations for the header
availableCell = 1 # This is used to determine the next cell to write onto, by counting up which cells are used with iteration and previous availableCell

workbook = xlrd.open_workbook(readBook)
worksheet = workbook.sheet_by_name(sheetname)
#This opens the file given and the sheet to be read

workbookWrite = xlsxwriter.Workbook(writeBook)
sheet = workbookWrite.add_worksheet('Sheet_1')
#This creates the file given and a sheet to be written onto

for header in ["Word", "everyFREQ", "PoS", "FREQcount", "SUBTLWF", "Zipf-value"]: # Creates the header
    sheet.write(0, headerCounter, header)
    headerCounter += 1

for k in range(1, rowNumber): # Main loop which retrieves, processes, and writes the data
    print(k)
    frq = retrData(k, 13)
    pos = retrData(k, 12)
    word = retrData(k, 0)
    totalFrqList = (retrData(k, 1))
    totalFrq = float(totalFrqList[0])
    gdict = correlateValues(pos, frq)
    finalDict = delName(gdict)
    availableCell = writeData(finalDict, word, availableCell, totalFrq)

workbookWrite.close()
#Saves the new excel file. --NOTE-- The file only works in xls form, for some reason :P