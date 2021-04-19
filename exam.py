import pandas as pd
import sys
from icalendar import Event, Calendar

def readExcel(name: str) -> pd.DataFrame:
    excel = pd.read_excel(name, engine='openpyxl', header=1).iloc[:,1:5]

    currentDate = None
    for row in range(len(excel)):
        if excel.isna().loc[row,'Date']:
            excel.at[row,'Date'] = currentDate
        else:
            currentDate = excel.loc[row,'Date']

    return excel.dropna()

def filter(raw: pd.DataFrame) -> pd.DataFrame:
    result = raw
    for filter

if __name__ == '__init__':
    # 创建Calendar类实例
    cal = Calendar()
    cal['CALSCALE'] = 'CALSCALE'
    cal['METHOD'] = 'METHOD'
    cal['X-WR-CALNAME'] = '考试表'
    cal['X-WR-TIMEZONE'] = 'Asia/Shanghai'

    excel = readExcel(sys.argv[1])

