import xlwings as xw
import os
currentPath=os.getcwd()
print(currentPath)
excelPath="D:\Benny_document\python\data"
excelFileName="9130i-pwr-chn.xlsx"
excelFilePath=excelPath +"\\" + excelFileName

wb=xw.Book(excelFilePath)
sheetNames=[]
for name in wb.sheets:
    sheetNames.append(name.name)
    #print(name.name)

#sheets.count show sheets total
print("Sheet number:" + str(wb.sheets.count))

sheet=wb.sheets[sheetNames[10]]
print(sheet)

#sheet.active()

print(sheet.used_range.last_cell.row)

print(sheet.used_range.last_cell.column)

#A1 (1,1) A2(2,1)
print(sheet.cells(3,4).value)
print("------------------")
#merge_area  return range <Range [9130i-pwr-chn.xlsx]5GHz -A M!$A$3:$C$6>
print(xw.Range('A3').merge_area)
print(xw.Range('B3').merge_area)
print(xw.Range('C3').merge_area)
print(xw.Range('D3').merge_area)
print(xw.Range('C3').merge_area)
mergeRange=str(xw.Range('C3').merge_area)
print(len(mergeRange))
print(mergeRange.find("$"))
mergeRange=mergeRange[37:46]
mergeRange=mergeRange.replace("$",'')
mergeRange=mergeRange.split(':')
print(mergeRange)

def searchMergeRange(location):
    merageRange=str(xw.Range(location).merge_area)
    merageRangeSize=len(merageRange)
    merageRangeStart=merageRange.find("$")
    merageRange=merageRange[merageRangeStart:merageRangeSize-1]
    merageRange=merageRange.replace("$",'')
    merageRange=merageRange.split(':')
    return merageRange

def showMergeRange(rangeList):
    rowsCharacter =rangeList[0][0]
    columnsCharacter=rangeList[1][0]
    rowsNumber=rangeList[0][1]
    columnsNumber=rangeList[1][1]
    if rowsCharacter == columnsCharacter:
        startLocation=rowsCharacter+rowsNumber
        endCharacter=rowsCharacter

    return rowsCharacter,columnsCharacter

location=searchMergeRange('D5')
print(location)
print(showMergeRange(location))
wb.close()