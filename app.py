import pandas as pd
import xlwings as xw
import os

# Remove the old result file if exists
if os.path.exists("result.xlsx"):
  os.remove("result.xlsx")

# Load file into pandas
data = pd.read_excel("C:/Users/sanketp60/Desktop/Fox_Trading_Internship_2/Task/NIFTY25JUN2010000PE.xlsx", sheet_name="NIFTY25JUN2010000PE", parse_dates = [['Date', 'Time']])

# Take date_time field as index for frame conversion
data.index = pd.to_datetime(data['Date_Time'])

def opcl(s, operation):
    try:
        if operation == 'open':
            return s[0]
        if operation == 'close':
            return s[-1]
    except:
        pass

# Timeframe conversion from 1 min to 15 min
data = data.resample('15T').agg({'Open ': lambda s: opcl(s, 'open'), 
'High ': lambda s: s.max(), 
'Low ': lambda s: s.min(), 
'Close ': lambda s: opcl(s, 'close'), 
'Volume': lambda s: s.sum()})

# Remove fields with zero volume
data = data[data['Low '].notna()]

# Export result data to excel file
data.to_excel("result.xlsx", sheet_name = 'resampled data')
print("Resampled data stored in file named: result.xlsx")

# Delete data to save memory
del data

# Open result workbook and add stats sheet
wb = xw.Book('result.xlsx')
reader = wb.sheets['resampled data']
stats = wb.sheets.add('stats')
stats.range('A1').value = 'Date'
stats.range('B1').value = 'Profit'

length  = reader.range('A2').end('down').row
data = reader.range('A2:F'+str(length)).value

'''
0 - Date_Time
1 - Open
2 - High
3 - Low
4 - Close
5 - Volume
'''

# Result calculator method for the day
def day_result(data):
    n = len(data[1])
    current = 1
    previous = 0
    short_pointer = -1
    short = False

    while current<n:
        if data[3][current]<data[3][previous]:
            short = True
            short_pointer = current
            break
        previous = current
        current += 1
    if short:
        while current<n:
            if data[2][short_pointer] < data[2][current]:
                return data[4][current] - data[1][0]
            current +=1
    return data[4][-1] - data[1][0]

# Add ending character for indication of End of the list
data.append('EOF')

# Set printer pointer
count = 2

while data[0]!='EOF':
    date = data[0][0].date()
    i = 0
    buffer = []
    try:
        while data[i][0].date() == date:
            buffer.append(data[i])
            i += 1
    except AttributeError:
        pass
    data = data[i:]
    stats.range('A'+str(count)).value = [date, day_result(list(map(list, zip(*buffer))))]
    count += 1

# Print Success
print("Result calculated successfully!")