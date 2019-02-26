from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import getpass as getpass
import sys
import xlrd
import time as ti
from datetime import time

t = 0.5

class Webpage:
    def __init__(self, url):
        self.url = url
        self.driver = webdriver.Chrome()
        self.driver.get(url)

    def sendValuetoXpath(self, xpath, val):
        elem = self.driver.find_element_by_xpath(xpath)
        elem.click()
        elem.clear()
        elem.send_keys(val)

    def sendEnter(self, xpath):
        self.driver.find_element_by_xpath(xpath).send_keys(Keys.ENTER)

    def clickPath(self, xpath):
        elem = self.driver.find_element_by_xpath(xpath)
        elem.click();


def importReservations(filename):
    with open(filename, 'r') as f:
        data = json.load(f)
    return data

def initiate(username, password):
    # Log in
    page = Webpage("http://mediaworks.ucr.edu/")
    page.sendValuetoXpath('//*[@id="username"]', username)
    page.sendValuetoXpath('//*[@id="password"]', password)

    print("Logging in...")
    page.clickPath('//*[@id="login"]/table/tbody/tr[4]/td/input[4]')

    return page

def submitReservations(page, data, equipment, locations):
    # New Order
    for i in range(0,len(locations[0])):
        page.clickPath('//*[@id="user-nav"]/li[1]/a')
        # Fill in data
        page.sendValuetoXpath('//*[@id="p_eventContact"]', data[0])
        page.sendValuetoXpath('//*[@id="p_eventContactEmail"]', data[1])
        page.sendValuetoXpath('//*[@id="p_eventName"]', data[2])


        # start time / meridiem
        page.sendValuetoXpath('//*[@id="p_eventStartTime"]', data[3])
        page.clickPath('//*[@id="p_eventStartTimeAMPM_chzn"]/a')
        page.sendValuetoXpath('//*[@id="p_eventStartTimeAMPM_chzn"]/div/div/input', data[4])
        page.sendEnter('//*[@id="p_eventStartTimeAMPM_chzn"]/div/div/input')
        ti.sleep(t)

        # end time / meridiem
        page.sendValuetoXpath('//*[@id="p_eventEndTime"]', data[5])
        page.clickPath('//*[@id="p_eventEndTimeAMPM_chzn"]/a')
        page.sendValuetoXpath('//*[@id="p_eventEndTimeAMPM_chzn"]/div/div/input', data[6])
        page.sendEnter('//*[@id="p_eventEndTimeAMPM_chzn"]/div/div/input')
        ti.sleep(t)

        # building
        page.clickPath('//*[@id="p_eventBuilding_chzn"]')
        page.sendValuetoXpath('//*[@id="p_eventBuilding_chzn"]/div/div/input', (locations[0])[i])
        page.sendEnter('//*[@id="p_eventBuilding_chzn"]/div/div/input')
        ti.sleep(t)

        # floor
        page.clickPath('//*[@id="p_eventFloor_chzn"]/a/div/b')
        page.sendValuetoXpath('//*[@id="p_eventFloor_chzn"]/div/div/input', (locations[1])[i])
        page.sendEnter('//*[@id="p_eventFloor_chzn"]/div/div/input')
        ti.sleep(t)

        # room
        page.clickPath('//*[@id="p_eventRoom_chzn"]/a/div/b')
        page.sendValuetoXpath('//*[@id="p_eventRoom_chzn"]/div/div/input', (locations[2])[i])
        page.sendEnter('//*[@id="p_eventRoom_chzn"]/div/div/input')
        ti.sleep(t)

        # fau
        page.sendValuetoXpath('//*[@id="fauList"]/div[2]/div/div[1]/input', (data[7])[:6])
        page.sendValuetoXpath('//*[@id="fauList"]/div[3]/div/div[1]/input', (data[7])[7:12])
        page.sendValuetoXpath('//*[@id="fauList"]/div[4]/div/div[1]/input', (data[7])[13:15])
        page.sendValuetoXpath('//*[@id="fauList"]/div[5]/div/div[1]/input', data[8])
        page.sendValuetoXpath('//*[@id="fauList"]/div[6]/div/div[1]/input', data[9])

        # Equipment
        x = 1
        for i in equipment:
            # click plus button
            page.clickPath('//*[@id="btn_addEquipment"]')
            # click dropdown arrow
            page.clickPath('//*[@id="equipment' + str(x) + '_chzn"]/a/div/b')
            # fill in equipment name
            page.sendValuetoXpath('//*[@id="equipment' + str(x) + '_chzn"]/div/div/input', i)
            page.sendEnter('//*[@id="equipment' + str(x) + '_chzn"]/div/div/input')
            x += 1

        # user select date(s)
        page.clickPath('//*[@id="date_maker"]/div[1]/label[1]')
        page.sendValuetoXpath('//*[@id="firstDate"]', data[10])
        page.sendValuetoXpath('//*[@id="untilEndDate"]', data[10])

        # request items
        page.clickPath('//*[@id="p_action_save"]')
        ti.sleep(t)

        # submit
        page.clickPath('//*[@id="p_action_submit"]')
        ti.sleep(1)
        page.driver.switch_to_alert().accept()
        print("[ Order Submitted ]")
        ti.sleep(1)
        page.driver.switch_to_alert().accept()


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

    # for x in range(0, len(buildings)):
        # print("Building: " + buildings[x].value + "\n" + "Floor: " + floors[x].value + "\n" + "Room: " + (str(rooms[x].value))[0:-2] + "\n")

    return (buildings, floors, rooms)


srcFile = 'test.xlsx'

data = staticData(srcFile)
equipment = equipment(srcFile)
locations = getRooms(srcFile)

print "-----------------"
print data
print equipment
print locations


user = raw_input("NetID: ")
password = getpass.getpass("Password: ")
p = initiate(user, password)
submitReservations(p, data, equipment, locations)
p.driver.close()
