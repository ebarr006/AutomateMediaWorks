import xlrd
import time as ti
import getpass as getpass
from datetime import time

def floatToTime(x):
    x = int(x * 24 * 3600)
    hour = x // 3600
    minu = (x%3600)//60
    if hour > 12:
        meridiem = "PM"
        hour -= 12
    else:
        meridiem = "AM"

    if minu == 0:
        minu = "00"
    x = str(hour) + ":" + str(minu)
    return (x, meridiem)

def staticData(filename):
    data = []
    wb = xlrd.open_workbook(filename)
    main = wb.sheet_by_index(0)
    headers = main.row_values(0)
    input = main.row_values(1)

    data.append(input[headers.index('Name')])
    data.append(int(input[headers.index('Number')]))
    data.append(input[headers.index('Event')])

    t = floatToTime(input[headers.index('Start')])
    data.append(t[0])
    data.append(t[1])

    t = floatToTime(input[headers.index('End')])
    data.append(t[0])
    data.append(t[1])

    data.append(input[headers.index('FAU')])
    data.append(input[headers.index('Cost Center')])
    data.append(input[headers.index('Project Code')])

    year = xlrd.xldate_as_tuple(input[headers.index('Date')],0)[0]
    month = xlrd.xldate_as_tuple(input[headers.index('Date')],0)[1]
    day = xlrd.xldate_as_tuple(input[headers.index('Date')],0)[2]

    data.append(str(month) + "/" + str(day) + "/" + str(year))
    return data

def equipment(filename):
    data = []
    wb = xlrd.open_workbook(filename)
    equip = wb.sheet_by_index(1)
    for item in equip.col_values(0):
        data.append(item)
    #delete header
    del data[0]
    return data

def getRooms(filename):
    buildings = []
    floors = []
    rooms = []
    wb = xlrd.open_workbook(filename)
    locations = wb.sheet_by_index(2)
    for x in range(1, locations.nrows):
        buildings.append(locations.cell(x, 0).value)
        floors.append(locations.cell(x, 1).value)
        rooms.append(int(locations.cell(x,2).value))
    return (buildings, floors, rooms)
