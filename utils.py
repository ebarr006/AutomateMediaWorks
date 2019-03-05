import xlrd
import json
import time as ti
import getpass as getpass
from datetime import time

Abbrev = {
    "BRNHL" : "Bourns Hall",
    "CHUNG" : "Winston Chung Hall",
    "HMNSS" : "Humanities Socsci",
    "INTN"  : "CHASS Int North",
    "INTS"  : "CHASS Int South",
    "MSE"   : "Materials Science & Engineering",
    "OLMH"  : "Olmsted Hall",
    "SKYE"  : "Skye",
    "SPR"   : "Sproul Hall",
    "UNLH"  : "UNIV Lec Hall",
    "WAT"   : "Watkins Hall"
}

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

    wb = xlrd.open_workbook("test.xlsx")
    locations = wb.sheet_by_index(2)

    for x in range(1, locations.nrows):
        spot = locations.cell(x,0).value.split(" ")
        buildings.append(Abbrev[spot[0]])
        roomNumber = spot[1].rstrip()
        first = roomNumber[0]
        if first.isalpha():
            first = roomNumber[1]
            # print "length: " + str(len(roomNumber[1:]))
            if len(roomNumber[1:]) == 3:
                rooms.append(roomNumber[0:1] + '0' + roomNumber[1:])
                # first = roomNumber[2]

        else:
            first = roomNumber[0]
            if len(roomNumber) == 3:
                rooms.append("0" + roomNumber)
            else:
                rooms.append(roomNumber)

        if first == '0':
            floors.append("Lower Level")
        elif first == '1':
            floors.append("First Floor")
        elif first == '2':
            floors.append("Second Floor")
        elif first == '3':
            floors.append("Third Floor")
        elif first == '4':
            floors.append("Fourth Floor")

        print buildings[x-1] + " " + floors[x-1] + " " + rooms[x-1]

    return (buildings, floors, rooms)
