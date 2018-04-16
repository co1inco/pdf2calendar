import gCalendar as cal

calendar = cal.gCalendar()

start = "2018-04-16T14:00"
end = "2018-04-16T15:00"
cid = "vve3ivkcdl11m87ub5rt7lpnrk@group.calendar.google.com"

calendar.createEvent(cid, start, end)
