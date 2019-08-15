from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import selenium.common.exceptions as seleniumExceptions
import selenium.webdriver.support.select
import traceback
import shutil
import datetime
import os
import argparse
import time
import csv

driver = webdriver.Chrome()
driver.implicitly_wait(3)


class Table:
    table = None #This should be an array

    def filter(self, criteria):
        pass
        #return filteredTable

    def __init__(self, tableID):
        pass

class CSVManagement:
    @staticmethod
    def importCSV(self, inputCSV):
        print("Importing items...")
        with open(inputCSV) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            outputArray = []
            for row in csv_reader:
                # serial = row[0]
                # if row[1]:
                #    host = row[1]
                # if row[2]:
                #    ritm = row[2]
                outputArray.append(row[0])
                print("Importing line: " + str(row))
                line_count += 1
            print(f'Imported {line_count} items')
            print("Result: " + str(outputArray))
            return outputArray


class SelfService:

    def __init__(self):
        self.url = 'https://ghsprod.service-now.com/nav_to.do?uri=%2Fcatalog_home.do%3Fsysparm_view%3Dcatalog_default'
        driver.get(self.url)

    class DesktopSupport:

        class AssetManagement:
            def __init__(self):
                pass

        class MassComputerUpdate:
            def __init__(self):
                pass


class RequestItem:
    state = None
    number = None
    item = None
    request = None
    requestedBy = None
    dueDate = None
    configurationItem = None
    opened = None
    openedBy = None
    updated = None
    stage = None
    quantity = None
    estimatedDelivery = None

    def __init__(self):
        pass


class Tasks:
    incidents = []
    catalogTasks = []
    requestItems = []

    def __init__(self):
        #List incidents, catalog tasks, and requestitems
        self.catalogTasks.append('testTask')


class CatalogTask:
    Number = None
    requestItem = None
    requestedBy = None
    requestedFor = None
    dueDate = None
    configurationItem = None
    openedBy = None
    opened = None
    kbArticle = None
    approval = None
    taskName = None
    state = None
    myPriority = None
    assignmentGroup = None
    assignedTo = None

    def __init__(self):
        pass


class Incident:
    pass


class Change:
    pass

class Computer:


    def __init__(self, hostname):
        self.hostname = hostname
        self.tasks = Tasks



class ServiceNow:
    #selfService = SelfService()

    @staticmethod
    def switchToMainFrame():
        driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))

    @staticmethod
    def switchToDefaultFrame():
        driver.switch_to.default_content()

    def search(self, item):
        self.switchToDefaultFrame()
        print("Searching for " + item)
        searchbox = driver.find_element_by_xpath("//*[@id='sysparm_search']")
        searchbox.clear()
        searchbox.click()
        searchbox.send_keys(item)
        searchbox.send_keys(Keys.RETURN)
        self.switchToMainFrame()


serviceNow = ServiceNow()
#hostnames = CSVManagement.importCSV('input.csv')

computer = Computer('dt123')
print(str(computer.tasks.catalogTasks))

