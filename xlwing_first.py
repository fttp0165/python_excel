import xlwings as xw
import os
currentPath=os.getcwd()
print(currentPath)
excel_path="D:\Benny_document\python\data"
excel_file_name="9130i-pwr-chn.xlsx"
excel_file_path = excel_path + "\\" + excel_file_name
current_path=os.getcwd()
print(current_path)
wb = xw.Book(excel_file_path)
app = xw.apps.active
sheet_names=[]
#for name in wb.sheets:
#    sheet_names.append(name.name)
    #print(name.name)

#sheets.count show sheets total
#print("Sheet number:" + str(wb.sheets.count))

#sheet=wb.sheets[sheetNames[10]]
#print(sheet)

#sheet.active()

#print(sheet.used_range.last_cell.row)

#print(sheet.used_range.last_cell.column)

#A1 (1,1) A2(2,1)
#print(sheet.cells(3,4).value)
print("------------------")
#merge_area  return range <Range [9130i-pwr-chn.xlsx]5GHz -A M!$A$3:$C$6>
#print(xw.Range('A3').merge_area)
#print(xw.Range('B3').merge_area)
#print(xw.Range('C3').merge_area)
#print(xw.Range('D3').merge_area)
#print(xw.Range('C3').merge_area)
#mergeRange=str(xw.Range('C3').merge_area)
#print(len(mergeRange))
#print(mergeRange.find("$"))
#mergeRange=mergeRange[mergeRange.find("$"):len(mergeRange)-1]
#mergeRange=mergeRange.replace("$",'')
#mergeRange=mergeRange.split(':')
#print(mergeRange)

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
    range_cell=[]
    if rowsCharacter == columnsCharacter:
       for loc_number in range(rowsNumber,columnsNumber):
           range_cell.append(rowsCharacter+loc_number)
    elif rowsNumber == columnsNumber:
        print("rowsCharacter:",rowsCharacter)
        print("columnsCharacter:",columnsCharacter)
        for loc_character in range(ord(rowsCharacter),ord(columnsCharacter)+1):
            print("loc_character:",loc_character)
            range_cell.append(chr(loc_character)+rowsNumber)
    else:
        print("rowsCharacter:",rowsCharacter)
        print("columnsCharacter:",columnsCharacter)
        print("rowsNumber:",rowsNumber)
        print("columnsNumber:",columnsNumber)
        for loc_character in range(ord(rowsCharacter),ord(columnsCharacter)+1):
            for loc_number in range(int(rowsNumber),int(columnsNumber)+1):
                range_cell.append(chr(loc_character)+str(loc_number))
    return range_cell

def main():
    location=searchMergeRange('A3')
    print(location)
    print(showMergeRange(location))
    #wb.close()   #?????????workBook
    app.quit()


if __name__ == "__main__":
   main()
