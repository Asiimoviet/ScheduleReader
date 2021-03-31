import datetime
import sys
import pandas as pd
from icalendar import Calendar, Event

timeFormat = '%Y%m%dT%H%M%S'

# 创建Calendar类实例
cal = Calendar()
cal['CALSCALE'] = 'CALSCALE'
cal['METHOD'] = 'METHOD'
cal['X-WR-CALNAME'] = '课程表'
cal['X-WR-TIMEZONE'] = 'Asia/Shanghai'

# 读取Excel课程表并导入到DataFrame中
#path = input("Please Input the path of the excel schedule: ")
excel = pd.read_excel(sys.argv[1])
keywords = pd.read_json(sys.argv[2]).loc[:,'classes']

# 处理周一的日期
monday = datetime.date.fromisoformat(pd.read_json(sys.argv[2]).loc[:,'monday'][0])
rowIndex = excel.iloc[:, 0].dropna().index

# 提取课程
def extractClass(timeCol: int, col: int, dayOffset: int):
    baseDate = monday + datetime.timedelta(days = dayOffset)
    print(baseDate.ctime())
    for periodCount in range(9):
        print('')
        print("======================================================================")
        print('Processing {}th Period'.format(periodCount + 1))

        rowStart = rowIndex[periodCount]
        rowEnd = rowIndex[periodCount + 1]
        print('Reading from Row {} to {}'.format(rowStart, rowEnd))

        timeRaw = excel.iloc[rowStart, timeCol]
        timeStart = datetime.datetime(baseDate.year, baseDate.month, baseDate.day, int(timeRaw[:timeRaw.find(':')]), int(timeRaw[timeRaw.find(':') + 1:timeRaw.find('-')]))
        timeEnd = datetime.datetime(baseDate.year, baseDate.month, baseDate.day, int(timeRaw[timeRaw.find('-') + 1:timeRaw.rfind(':')]), int(timeRaw[timeRaw.rfind(':')+1:]))
        print('Time: {}'.format(timeRaw))
        print('Time Start: {}'.format(timeStart.strftime(timeFormat)))
        print('Time End: {}'.format(timeEnd.strftime(timeFormat)))
        print("======================================================================")

        periodsRaw = excel.iloc[rowStart:rowEnd, col].dropna()
        periods = []
        choice = -1
        
        print('0: I don\'t have a class at this period')
        for index in range(len(periodsRaw)):
            if periodsRaw.iloc[index] == '/': continue
            if periodsRaw.iloc[index].find(' - S') != -1:
                periods.append({
                    'summary': periodsRaw.iloc[index][:periodsRaw.iloc[index].rfind(' - ')],
                    'location': periodsRaw.iloc[index][periodsRaw.iloc[index].rfind(' - ') + 3:]
                })
            else:
                periods.append({
                    'summary': periodsRaw.iloc[index],
                    'location': ''
                })

            print('{}: {}'.format(len(periods), periods[len(periods) - 1]['summary']))

            for keyword in keywords:
                if periods[len(periods) - 1]['summary'].find(keyword) != -1:
                    choice = len(periods) - 1
                    break

        print('')

        if choice == -1: 
            print('No class found matching, skipping')
            continue

        print('Find matching: {}'.format(periods[choice]['summary']))

        period = periods[choice]
        event = Event()
        event['summary'] = period['summary']
        event['location'] = period['location']
        event['dtstart'] = timeStart.strftime(timeFormat)
        event['dtend'] = timeEnd.strftime(timeFormat)
        cal.add_component(event)


extractClass(1, 2, 0)
for day in range(4):
    extractClass(3, day + 4, day + 1)
    
# 导出文件
exportFile = open('result.ics', 'w+')
exportFile.write(cal.to_ical().decode('UTF-8'))
exportFile.close()