
try:
    import xlrd
except:
    pass

import re
import datetime

from tkinter import *
from tkinter import messagebox
import os

dateCol = 1
timeRow = 0

global monthList
monthList = ['Jan', 'Feb', 'Mar',
         'Apr', 'Mai', 'Jun',
         'Jul', 'Aug', 'Sep',
         'Okt', 'Nov', 'Dez']

def errorAndExit(message):
    messagebox.showerror("Error", message)
    os._exit(0)

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
        
            currentCell = self.getCell(page, timeRow, i)
#            print(str(len(currentCell)) + " " + currentCell)
            
            if currentCell.find('empty') != -1:
                pass
            elif currentCell.find('number') != -1:
                time = currentCell[7:] + "0 - " + str(float(currentCell[7:])+1.30) + "0 Uhr"
                
                print("Time: " + time + " -Repaired-")
                timeCols.append(i)
                times.append(time)
            elif 16 <= len(currentCell) <= 17:
                print("Time: " + currentCell)
                timeCols.append(i)
                times.append(currentCell)
                
        print(str(timeCols))
        return timeCols, times


    def getDay(self, page, row, timeCols, timeStr, dateCol = 1):

        classes = []
#        print(timeCols)#----
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
                try:    
                    time2 = time[2]
                except:
                    errorAndExit("Incorect time\n Please check the\n first Row of each sheet")
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

        if page.ncols < 5:
            return classes

        for i in range(page.nrows):
            currentCell = self.getCell(page, i, dateCol)
            if currentCell.find('empty') == -1:
                day = self.getDay(page, i, timeCols, timeStr, dateCol)

                for i in day:
                    classes.append(i)
        return classes


    def getAllPages(self):

        classes = []
        for i in range(self.book.nsheets):
            sh = self.book.sheet_by_index(i)
            page = self.getPage(sh)
            for j in page:
                classes.append(j)
            print("Page (%i) ready" % i)

        return classes

def getFileContent(filename):
    f = open(filename, 'r', encoding="utf-8")
    content = f.readlines()
    
    classes = []
    for i in content:
        tmp = re.split(r'\t+', i.rstrip('\t'))
#        tmp = tmp[:-1]
#        print(i)
        classes.append(tmp)

    f.close()

    return classes

def writeToFile(filename, classes):
        
    f = open(filename, 'w', encoding="utf-8")
    for i in classes:

        for j in i:
            if type(j).__name__ == 'str':
                f.write(j + '\t')
            if type(j).__name__ == 'list':
                for k in j:
                    f.write(k + ',')
                f.write('\t')

        f.write("\n")

    f.close()

    

if __name__ == '__main__':
       
    read = readXlsx("Stundenplan_WS 2018-19_ELM_1.xlsx")
    
    classes = read.getAllPages()

    writeToFile("classes.txt", classes)

    g = getFileContent("classes.txt")
    for i in g:
        print(i)



    

    
    
