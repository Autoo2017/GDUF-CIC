from datetime import datetime, timedelta
from icalendar import Calendar, Event, Alarm

import pandas as pd


def read(filenames):


    cal = Calendar()
    cal.add('prodid', '-//My calendar product//mxm.dk//')
    cal.add('version', '2.0')
    path = 'user/' + filenames
    df = pd.read_excel(path)

    # cal = Calendar()
    # cal.add('prodid', '-//My calendar product//mxm.dk//')
    # cal.add('version', '2.0')
    # path = 'user/example.xls'
    # df = pd.read_excel(path)

    start = [datetime(2023,2,20,8,15,0),
            datetime(2023,2,20,10,0,0),
            datetime(2023,2,20,11,35,0),
            datetime(2023,2,20,14,0,0),
            datetime(2023,2,20,15,35,0),
            datetime(2023,2,20,16,30,0),
            datetime(2023,2,20,19,0,0)]

    #横排叫行（row），竖排为列(column)
    def find_row():
        for i in range(0,df.shape[1]):
            for j in range(0,df.shape[0]):
                check = df.iloc[j,i]
                if check == "星期一":
                    return i

    row =  find_row()

    def find_column():

        for i in range( 0,df.shape[1] ):
            a = df.iloc[row, i]
            if a == "星期一":
                return i


    column = find_column()

    for k in range(17):



        for j in range(column ,df.shape[1]):

                for i in range(row +1,df.shape[0] -1):

                    isMorN = df.iloc[i,0]
                    if isMorN == 'M'  or isMorN =='N':
                        end = start[i - 2] + timedelta(minutes=45)
                    else:
                        end = start[i - 2] + timedelta(minutes=90)

                    event = Event()
                    res = df.iloc[i, j].split("\n")

                    try:
                        title = res[1]
                        teacher = res[2]
                        location = res[4]
                    except:
                        start[i - 2] = start[i - 2] + timedelta(days=1)
                        continue



                    event.add('summary', title)
                    event.add('location', location)
                    event.add('description', teacher)
                    event.add('dtstart', start[i-2])
                    event.add('dtend', end)
                    event.add('dtstamp', datetime.utcnow())
                    event.add = ('uid', k * 100 + j * 10 + i * 1)


                    alarm = Alarm()
                    alarm.add('trigger', start[i-2] + timedelta(minutes = -30))
                    alarm.add('action', 'DISPLAY')
                    alarm.add('description', "Lesson Reminder")
                    event.add_component(alarm)

                    cal.add_component(event)

                    start[i-2] = start[i-2] + timedelta(days = 1)

    start[i-2] = start[i-2] + timedelta(weeks = 1)

    with open(path+".ics", "wb") as f:
        f.write(cal.to_ical())
