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


class CatalogTask:
    number = None
    requestItem = None
    request = None
    item = None
    requestedFor = None
    contactNumber = None
    alternateNumber = None
    shortDesciption = None
    description = None
    state = None
    assignmentGroup = None
    assignedTo = None
    taskDueDate = None
    configurationItem = None

    def __init__(self):
        switchToDefaultFrame()
        switchToMainFrame()
        self.number = driver.find_element_by_id("sc_task.number").get_attribute("value")

driver.get("https://prismahealth.service-now.com")
#Login
driver.find_element_by_xpath("//*[@id='maincontent']/tbody/tr[4]/td[2]").click()

#Go to a task
driver.get("https://prismahealth.service-now.com/nav_to.do?uri=task.do?sys_id=b537bb9e1bbff308f15a4005bd4bcb8b")

task = CatalogTask()

print(task.number)

