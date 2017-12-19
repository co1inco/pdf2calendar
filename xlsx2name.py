import xlrd
import re
import datetime

dateCol = 1
timeRow = 0

global monthList
monthList = ['Jan', 'Feb', 'Mar',
         'Apr', 'Mai', 'Jun',
         'Jul', 'Aug', 'Sep',
         'Okt', 'Nov', 'Dez']

class readXlsx():
    def __init__(self, filename):

        self.book = xlrd.open_workbook(filename)


    def getCell(self, page, row, col):
        cell = str(page.row(row)[col])
        if cell.find('text') > -1:
            cell = cell[6:-1]
        return cell


    def getTime(self, page, timeRow = 0):
        timeCols = []
        times    = []   

        for i in range(1, page.ncols):
        
            currentCel = self.getCell(page, timeRow, i)
            if currentCel.find('empty') == -1:
                timeCols.append(i)
                times.append(currentCel)

        return timeCols, times


    def getDay(self, page, row, timeCols, timeStr, dateCol = 1):

        classes = []

        for timeIndex, timeCol in enumerate(timeCols):
        
            currentCel = self.getCell(page, row, timeCol)
            if currentCel.find('empty') == -1:

                #2017-05-28T17:00:00-00:00
                date = self.getCell(page, row, dateCol)
                year = str(datetime.datetime.now())[0:4]
                month = date[4:7]
                day = date[0:2]
                for i, j in enumerate(monthList):
#                    print(i+1, j)
                    if month == j:
                        month = i + 1
                        break
                date = year + "-" + str(month) + "-" + str(day)

                time = timeStr[timeIndex]
                time = time.split()
                time1 = time[0]
                if len(time1) == 4:
                    time1 = "0" + time1
                time2 = time[2]
                if len(time2) == 3:
                    time2 = "0" + time2
                time = str(time1) + "-" + str(time2)

                clas = []
                clas.append(date)
                clas.append(time)
                clas.append(re.sub(' +', ' ', currentCel))

                currentCol = timeCols[timeIndex] + 1
            
                try:
                    nextClassCol = timeCols[timeIndex+1]
                except IndexError:
                    nextClassCol = page.ncols
                while currentCol < nextClassCol:
                    currentCel = self.getCell(page, row, currentCol)
                    if currentCel.find('empty') == -1:
                        clas.append(currentCel)
                    currentCol = currentCol + 1

                classes.append(clas)

        return classes


    def getPage(self, page, dateCol = 1):

        timeCols, timeStr = self.getTime(page)
        classes = []

        for i in range(page.nrows):
            currentCell = self.getCell(page, i, dateCol)
            if currentCell.find('empty') == -1:
                day = self.getDay(page, i, timeCols, timeStr, dateCol)

                for i in day:
                    classes.append(i)

        return classes


    def getAllPages(self):

        classes = []
        for i in range(self.book.nsheets-1):
            sh = self.book.sheet_by_index(i)
            page = self.getPage(sh)
            for j in page:
                classes.append(j)

        return classes

def getFileContent(filename):
    f = open('classes.txt', 'r', encoding="utf-8")
    content = f.readlines()
    
    classes = []
    for i in content:
        tmp = re.split(r'\t+', i.rstrip('\t'))
        tmp = tmp[:-1]
        classes.append(tmp)

    return classes

    

if __name__ == '__main__':
    
#    book = xlrd.open_workbook("Stundenplan_WS 2017-18_ELM 1 ilovepdf og.xlsx")
    
    read = readXlsx("Stundenplan_WS 2017-18_ELM 1.xlsx")

    f = open('classes.txt', 'w', encoding="utf-8")

    classes = read.getAllPages()
    
    # [:9] date | [11:32] time | [34:] name
    for i in classes:

        for j in i:
            if type(j).__name__ == 'str':
                f.write(j + '\t')
            if type(j).__name__ == 'list':
                for k in j:
                    f.write(k + ',')
                f.write('\t')

        f.write("\n")
#        print(i)

    f.close()

    g = getFileContent("classes.txt")
    for i in g:
        print(i)



    

    
    
