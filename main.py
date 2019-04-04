from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils import *

t = 0.5

class Webpage:
    def __init__(self):
        self.user = raw_input("NetID: ")
        self.password = getpass.getpass("Password: ")
        self.driver = webdriver.Chrome()
        self.driver.get("http://mediaworks.ucr.edu/")
        self.sendValuetoXpath('//*[@id="username"]', self.user)
        self.sendValuetoXpath('//*[@id="password"]', self.password)
        self.clickPath('//*[@id="login"]/table/tbody/tr[4]/td/input[4]')
        print("Logging in...")

    def sendValuetoXpath(self, xpath, val):
        elem = self.driver.find_element_by_xpath(xpath)
        elem.send_keys(val)

    def sendEnter(self, xpath):
        self.driver.find_element_by_xpath(xpath).send_keys(Keys.ENTER)

    def clickPath(self, xpath):
        elem = self.driver.find_element_by_xpath(xpath)
        elem.click();

    def submitReservations(self, data, equipment, locations):
        # New Order
        for i in range(0,len(locations[0])):
            self.clickPath('//*[@id="user-nav"]/li[1]/a')
            # Fill in data
            self.sendValuetoXpath('//*[@id="p_eventContact"]', data[0])
            self.sendValuetoXpath('//*[@id="p_eventContactEmail"]', data[1])
            self.sendValuetoXpath('//*[@id="p_eventName"]', data[2])

            # start time / meridiem
            self.sendValuetoXpath('//*[@id="p_eventStartTime"]', data[3])
            self.clickPath('//*[@id="p_eventStartTimeAMPM_chzn"]/a')
            self.sendValuetoXpath('//*[@id="p_eventStartTimeAMPM_chzn"]/div/div/input', data[4])
            self.sendEnter('//*[@id="p_eventStartTimeAMPM_chzn"]/div/div/input')
            ti.sleep(t)

            # end time / meridiem
            self.sendValuetoXpath('//*[@id="p_eventEndTime"]', data[5])
            self.clickPath('//*[@id="p_eventEndTimeAMPM_chzn"]/a')
            self.sendValuetoXpath('//*[@id="p_eventEndTimeAMPM_chzn"]/div/div/input', data[6])
            self.sendEnter('//*[@id="p_eventEndTimeAMPM_chzn"]/div/div/input')
            ti.sleep(t)

            # building
            self.clickPath('//*[@id="p_eventBuilding_chzn"]')
            self.sendValuetoXpath('//*[@id="p_eventBuilding_chzn"]/div/div/input', (locations[0])[i])
            self.sendEnter('//*[@id="p_eventBuilding_chzn"]/div/div/input')
            ti.sleep(t)

            # floor
            self.clickPath('//*[@id="p_eventFloor_chzn"]/a/div/b')
            self.sendValuetoXpath('//*[@id="p_eventFloor_chzn"]/div/div/input', (locations[1])[i])
            self.sendEnter('//*[@id="p_eventFloor_chzn"]/div/div/input')
            ti.sleep(t)

            # room
            self.clickPath('//*[@id="p_eventRoom_chzn"]/a/div/b')
            self.sendValuetoXpath('//*[@id="p_eventRoom_chzn"]/div/div/input', (locations[2])[i])
            self.sendEnter('//*[@id="p_eventRoom_chzn"]/div/div/input')
            ti.sleep(t)

            # fau
            self.sendValuetoXpath('//*[@id="fauList"]/div[2]/div/div[1]/input', (data[7])[:6])
            self.sendValuetoXpath('//*[@id="fauList"]/div[3]/div/div[1]/input', (data[7])[7:12])
            self.sendValuetoXpath('//*[@id="fauList"]/div[4]/div/div[1]/input', (data[7])[13:15])
            self.sendValuetoXpath('//*[@id="fauList"]/div[5]/div/div[1]/input', data[8])
            self.sendValuetoXpath('//*[@id="fauList"]/div[6]/div/div[1]/input', data[9])

            # Equipment
            x = 1
            for i in equipment:
                print i
                # click plus button
                self.clickPath('//*[@id="btn_addEquipment"]')
                # click dropdown arrow
                self.clickPath('//*[@id="equipment' + str(x) + '_chzn"]/a/div/b')
                # fill in equipment name
                self.sendValuetoXpath('//*[@id="equipment' + str(x) + '_chzn"]/div/div/input', i)
                self.sendEnter('//*[@id="equipment' + str(x) + '_chzn"]/div/div/input')
                x += 1

            # user select date(s)
            self.clickPath('//*[@id="date_maker"]/div[1]/label[1]')
            self.sendValuetoXpath('//*[@id="firstDate"]', data[10])
            self.sendValuetoXpath('//*[@id="untilEndDate"]', data[10])

            # # request items
            self.clickPath('//*[@id="p_action_save"]')
            ti.sleep(t)

            # #submit
            page.clickPath('//*[@id="p_action_submit"]')
            ti.sleep(2)
            page.driver.switch_to_alert().accept()
            print("[ Order Submitted ]")
            ti.sleep(2)
            page.driver.switch_to_alert().accept()


if __name__ == "__main__":
    srcFile = 'test.xlsx'
    data = staticData(srcFile)
    equipment = equipment(srcFile)
    locations = getRooms(srcFile)
    page = Webpage()
    page.submitReservations(data, equipment, locations)
