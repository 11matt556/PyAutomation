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

class CatalogTasksPage:

    def __init__(self):
        pass

    def findRequestItem(self,cond):
        print(str(cond))
        #Create table. Search and filter table
        req = RequestItem()
        return req

class IncidentsPage:
    pass


class RequestedItemsPage:
    pass

class SearchPage:

    def clickCatalogTasks(self):
        cat = CatalogTasksPage()
        return cat

    def clickIncidents(self):
        incident = IncidentsPage()
        return incident

    def clickRequestedItems(self):
        requestedItems = RequestedItemsPage()
        return requestedItems

class ServiceNow:

    def searchFor(self, string): # Need to use composition or inheritance so everything can access global ServiceNow functionality
        print("Searching for " + string)
        searchbox = driver.find_element_by_xpath("//*[@id='sysparm_search']")
        searchbox.clear()
        searchbox.click()
        searchbox.send_keys(string)
        searchbox.send_keys(Keys.RETURN)
        results = SearchPage()
        return results

    class SelfService:
        class DesktopSupport:
            class AssetManagement:
                pass

            class MassComputerUpdate:
                pass

class RequestItem:
    state = None
    number = None
    item = 'yo'
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

serviceNow = ServiceNow()
serviceNow.searchFor('123').clickCatalogTasks().findRequestItem(['conditions'])
