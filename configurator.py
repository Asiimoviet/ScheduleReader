import sys
import json
import pandas as pd 

config = {
    'class': '',
    'selected': []
}

freePeriods = []

classes = ['ACT 1', 'ACT 2', 'ACT 3', 'SAT', 'TOEFL 2', 'TOEFL 4']

print('======================================================')
print('Selected Your Class')
print('======================================================')

for index in range(len(classes)):
    print('{}: {}'.format(index + 1, classes[index]))

choice = int(input('Please Input the number before the class you are in: '))
print('Selected Class: {}'.format(classes[choice - 1]))
config['class'] = classes[choice - 1]

print('Loading Schedule')
excel = pd.read_excel(sys.argv[1])

def swipeRow(row: int):
    rowList = excel.iloc[:, 0].dropna().index
    for classNum in range(9):
        rowStart = rowList[classNum]
        rowEnd = rowList[classNum + 1] - 1
        periods = excel.iloc[rowStart:rowEnd, row].dropna()
        found = False
        for period in periods:
            if period == '/': continue
            if period.find(config['class']) != -1: 
                found = True
                break
            for selected in config['selected']:
                if period.find(selected) != -1:
                    found = True
                    break
            if found: break

        if found: continue
        if len(periods) == 0: continue

        isFree = False
        
        for free in freePeriods:
            if periods.reset_index(drop=1).equals(free.reset_index(drop=1)):
                isFree = True
                break
        
        if isFree: continue

        print('\nSelected the subject you have chosen: ')
        print('0: I don\'t have a class at this period')
        for index in range(len(periods)):
            print('{}: {}'.format(index + 1, periods.iloc[index]))
        choice = int(input('Please input your choice: '))

        if choice == 0: 
            freePeriods.append(periods)
            continue 

        config['selected'].append(periods.iloc[choice - 1])
        
    
swipeRow(2)
for day in range(4):
    swipeRow(day + 4) 

print('Exporting...')
exportFile = open('config.json', 'w+')
exportFile.write(json.dumps(config))
exportFile.close()