import xlrd
import re
import xlwt

count = 0

def main():
# Open an Excel file
    sourceBook = xlrd.open_workbook("source.xls")
    targetBook = xlrd.open_workbook("target.xlsx")
    targetFile = xlwt.Workbook(encoding='utf-8')
    sheet = targetFile.add_sheet("result")
    for i in sourceBook.sheets():
        for rowIndex in range(i.nrows):
            if is2017(i.cell(rowIndex,10),sourceBook):
                contractIndex = i.cell(rowIndex,1).value.replace("-","").replace(" ","")

                if not isInTarget(contractIndex,targetBook):
                    with open("result.txt",'a') as f:
                        f.write(contractIndex+'\n\n')
                    global count
                    sheet.write(count,1,i.cell(rowIndex,1).value)
                    sheet.write(count,2,i.cell(rowIndex,4).value)
                    sheet.write(count,3,i.cell(rowIndex,5).value)
                    sheet.write(count,4,i.cell(rowIndex,8).value)
                    count +=1
    targetFile.save("result.xls")


def isInTarget(sourceStr:str,targetBook):
    sourceResult = re.findall(r"[续]?\d+", sourceStr)
    targetSheet = targetBook.sheet_by_name("新增集采清单")
    flag = False
    for rowIndex in range(targetSheet.nrows):
        targetStr = targetSheet.cell_value(rowIndex,1).replace("-","").replace(" ","")
        targetResult = re.findall(r"[续]?\d+",targetStr)
        if sourceResult==targetResult:
            flag = True
    return flag

def is2017(cellData,book):
    if cellData.ctype == 3: # When the data is "date"
        year,*other = xlrd.xldate_as_tuple(cellData.value,book.datemode)
        return year == 2017
    elif cellData.ctype == 1:  # When the data is "text"
        value = cellData.value
        result = True if value[:4]=="2017" or value[:2] == "17" else False
        return result
    else:
        return False


if __name__ == '__main__':
    main()