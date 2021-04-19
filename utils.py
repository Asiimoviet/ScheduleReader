import datetime
from dateutil import tz
import pandas as pd
import openpyxl

GOOGLE_CALENDAR_FORMATTER = '%Y%m%dT%H%M%SZ'

def dumpScheduleExcel(filename: str) -> pd.DataFrame:
    excel = pd.read_excel(filename, engine='openpyxl')

    for day in range(5):
        for period in range(9):
            

def dumpExamExcel(filename: str) -> pd.DataFrame:
    excel = pd.read_excel(filename, engine='openpyxl', header=1).iloc[:,1:5]

    currentDate = None
    for row in range(len(excel)):
        if excel.isna().loc[row,'Date']:
            excel.at[row,'Date'] = currentDate
        else:
            currentDate = excel.loc[row,'Date']

    return excel.dropna()

def dumpDate(rawDate: str):
    return datetime.date(
        year=datetime.date.today().year,
        month=int(rawDate[:rawDate.find('.')]),
        day=int(rawDate[rawDate.find('.') + 1:])
    )

def dumpDateTime(rawTime: str, rawDate: str, timeZone: datetime.tzinfo = tz.gettz('Asia/Shanghai')):
    date = dumpDate(rawDate)
    time = datetime.datetime(
        year=date.year,
        month=date.month,
        day=date.day,
        hour=int(rawTime[:rawTime.find(':')]),
        minute=int(rawTime[rawTime.find(':') + 1:]),
        tzinfo=timeZone
    )
    return time.astimezone(tz.UTC)

def dumpDateTimePeriod(rawPeriod: str, date: datetime.date):
    rawStartTime = rawPeriod[:rawPeriod.find('-')]
    rawEndTime = rawPeriod[rawPeriod.find('-') + 1:]

    return dumpDateTime(rawStartTime, date), dumpDateTime(rawEndTime, date)

if __name__ == '__main__':
    