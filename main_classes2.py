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

def switchToMainFrame():
    driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))


def switchToDefaultFrame():
    driver.switch_to.default_content()


class Notes:
    @staticmethod
    def setAdditionalComments(self, commentString):
        print("Typing " + str(commentString))
        commentArea = driver.find_element_by_xpath("//textarea[@id='activity-stream-comments-textarea']")
        # Sometimes the comment box area is collapsed by default, so we must click the button to expand it first
        try:
            commentArea.send_keys(commentString)
        except seleniumExceptions.ElementNotInteractableException:
            driver.find_element_by_xpath("//span[2]/span/nav/div/div[2]/span/button").click()
            commentArea.send_keys(commentString)


class Details:
    actualStart = None
    sequence = None
    serviceType = None
    taskName = None
    actualEnd = None
    location = None
    locationNotes = None

    def __init__(self):
        self.taskName = driver.find_element_by_id("sc_task.u_task_name").get_attribute("value")


class CatalogTask:
    notesTab = Notes()
    detailsTab = Details()
    variablesTab = None

    def __init__(self):
        switchToDefaultFrame()
        switchToMainFrame()
        self.number = driver.find_element_by_id("sc_task.number").get_attribute("value")
        self.requestItem = driver.find_element_by_id("sys_display.sc_task.request_item").get_attribute("value")
        self.shortDescription = driver.find_element_by_id("sc_task.short_description").get_attribute("value")
        self.state = selenium.webdriver.support.select.Select(driver.find_element_by_id('sc_task.state'))
        self.assignedTo = driver.find_element_by_id('sys_display.sc_task.assigned_to').get_attribute("value")

class VariablesRestock:
    pass

class VariablesRepair:
    pass


class RestockTask(CatalogTask):

    def __init__(self):
        CatalogTask.__init__(self)
        self.variablesTab = VariablesRestock()

    pass

class RepairTask(CatalogTask):
    #Stuff for repair
    pass

driver.get("https://prismahealth.service-now.com")
#Login
driver.find_element_by_xpath("//*[@id='maincontent']/tbody/tr[4]/td[2]").click()

#Go to a task
driver.get("https://prismahealth.service-now.com/nav_to.do?uri=task.do?sys_id=b537bb9e1bbff308f15a4005bd4bcb8b")

task = CatalogTask()

print(task.state)
print()
