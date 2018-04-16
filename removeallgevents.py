from gCalendar import *


calendar = gCalendar()

cList = calendar.getCalendarList()
for i, j in enumerate(cList):
    print(str(i) + ": " + j['summary'])

cid = cList[int(input("Type number of calendar: "))]['id']
print(cid)

x = calendar.getEventList(cid)

j = len(x)

for i in x:
    print(str(j) + " : " + i['id'])
    calendar.delEvent(cid, i['id'])
    j = j - 1

input("Ready")
