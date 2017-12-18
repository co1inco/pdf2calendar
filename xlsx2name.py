import xlrd
import re

dateCol = 1
timeRow = 0

def getCell(page, row, col):
    cell = str(page.row(row)[col])
    if cell.find('text') > -1:
        cell = cell[6:-1]
    return cell

def getTime(page, timeRow = 0):
    timeCols = []
    times    = []   

    for i in range(1, page.ncols):
        
        currentCel = getCell(page, timeRow, i)
        if currentCel.find('empty') == -1:
            timeCols.append(i)
            times.append(currentCel)

    return timeCols, times


def getDay(page, row, timeCols, timeStr, dateCol = 1):

    classes = []

    for timeIndex, timeCol in enumerate(timeCols):
        
        currentCel = getCell(page, row, timeCol)
        if currentCel.find('empty') == -1:

            date = getCell(page, row, dateCol)

            clas = []
            clas.append(date)
            clas.append(timeStr[timeIndex])
            clas.append(re.sub(' +', ' ', currentCel))

            currentCol = timeCols[timeIndex] + 1
            
            try:
                nextClassCol = timeCols[timeIndex+1]
            except IndexError:
                nextClassCol = page.ncols
            while currentCol < nextClassCol:
                currentCel = getCell(page, row, currentCol)
                if currentCel.find('empty') == -1:
                    clas.append(currentCel)
                currentCol = currentCol + 1

            classes.append(clas)

    return classes


def getPage(page, dateCol = 1):

    timeCols, timeStr = getTime(page)
    classes = []

    for i in range(page.nrows):
        currentCell = getCell(page, i, dateCol)
        if currentCell.find('empty') == -1:
            day = getDay(page, i, timeCols, timeStr, dateCol)

            for i in day:
                classes.append(i)

    return classes

def getAllPages(book):

    classes = []
    for i in range(book.nsheets-1):
        sh = book.sheet_by_index(i)
        page = getPage(sh)
        for j in page:
            classes.append(j)

    return classes
    

if __name__ == '__main__':
    
    book = xlrd.open_workbook("Stundenplan_WS 2017-18_ELM 1 ilovepdf og.xlsx")
    sh = book.sheet_by_index(2)

#    x, y = getTime(sh)

#    z = getDay(sh, 1, x, y)
    f = open('classes.txt', 'w', encoding="utf-8")

    classes = getAllPages(book)
    
    for i in classes:

        for j in i:
            f.write(j + '\t;\t')
        f.write("\n")
        
        print(i)
    f.close()

    
    
