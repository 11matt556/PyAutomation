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
    notesTab = None
    detailsTab = None
    variablesTab = None

    def __init__(self):
        switchToDefaultFrame()
        switchToMainFrame()
        self.notesTab = Notes()
        self.detailsTab = Details()
        self.number = driver.find_element_by_id("sc_task.number").get_attribute("value")
        self.requestItem = driver.find_element_by_id("sys_display.sc_task.request_item").get_attribute("value")
        self.shortDescription = driver.find_element_by_id("sc_task.short_description").get_attribute("value")
        self.state = selenium.webdriver.support.select.Select(driver.find_element_by_id('sc_task.state'))
        self.assignedTo = driver.find_element_by_id('sys_display.sc_task.assigned_to').get_attribute("value")

    def get_state(self):
        return self.state

    def set_state(self, state):
        print("Setting ticket state to " + state)
        driver.find_element_by_id('sc_task.state').find_element_by_xpath("//option[. = '" + state + "']").click()
        self.state = state


class RestockVariables:
    pass

class RepairVariables:
    pass


class RestockTask(CatalogTask):
    def __init__(self):
        CatalogTask.__init__(self)
        self.variablesTab = RestockVariables()

class RepairTask(CatalogTask):
    #Stuff for repair
    pass


class ServiceNow:
    @staticmethod
    def tickets_assigned_to_tim():
        driver.get("https://prismahealth.service-now.com/nav_to.do?uri=%2Ftask_list.do%3Fsysparm_fixed_query%3D%26sysparm_query%3Dactive%253Dtrue%255Eassigned_to%253D78ae421e1b4fb700f15a4005bd4bcb8d%26sysparm_clear_stack%3Dtrue")

    @staticmethod
    def device_management():
        driver.get("https://prismahealth.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D9b8d095f4f9e7b0416d5334d0210c7b4%26sysparm_link_parent%3D7edd9ad21b9f734033eb85507e4bcb98%26sysparm_catalog%3D816d92d21b9f734033eb85507e4bcbf6%26sysparm_catalog_view%3Dcatalog_desktop_support")

    @staticmethod
    def homepage():
        driver.get("https://prismahealth.service-now.com")


class Table:
    table = None
    table_head = None
    table_body = None
    rows_body = None
    rows_head = None

    def __init__(self, table_id):
        switchToDefaultFrame()
        switchToMainFrame()
        print("Loading table")
        self.table = driver.find_element_by_id(table_id)
        self.table_head = self.table.find_element_by_tag_name("thead")
        self.table_body = self.table.find_element_by_tag_name("tbody")

        elements = []
        self.rows_body = self.table_body.find_elements_by_tag_name("tr")
        self.rows_head = self.table_head.find_elements_by_tag_name("tr")

    def get_row_len_body(self):
        return len(self.rows_body)

    def get_row_len_head(self):
        return len(self.rows_head)

    def get_col_len_body(self):
        return len(self.rows_body[0].find_elements_by_tag_name('td')) - 1

    def get_col_len_head(self):
        return len(self.rows_head[0].find_elements_by_tag_name('th')) - 1

    def get_row_body(self, r):
        row_items = self.rows_body[r].find_elements_by_tag_name("td")
        print("Getting row " + str(r))
        for item in row_items:
            print(item.text)
        return row_items

    def get_row_head(self, r):
        row_items = self.rows_head[r].find_elements_by_tag_name("th")
        print("Getting row " + str(r))
        for item in row_items:
            print(item.get_attribute("glide_label"))
        return row_items

    def get_cell_body(self, r, c):
        cell = self.get_row_body(r)[c]
        print("row:" + str(r) + " col:" + str(c) + " is " + cell.text)
        return cell

    def get_cell_head(self, r, c):
        cell = self.get_row_head(r)[c]
        print("row:" + str(r) + " col:" + str(c) + " is " + cell.get_attribute("glide_label"))
        return cell

    def get_col_name(self, c):
        return self.get_cell_head(0, c).get_attribute("glide_label")


ServiceNow.homepage()

#Login
driver.find_element_by_xpath("//*[@id='maincontent']/tbody/tr[4]/td[2]").click()

#Go to a task
#driver.get("https://prismahealth.service-now.com/nav_to.do?uri=task.do?sys_id=b537bb9e1bbff308f15a4005bd4bcb8b")

#Go to tickets assigned to Tim"

ServiceNow.tickets_assigned_to_tim()

assigned = Table("task_table")
print(assigned.get_row_body(2))
print(assigned.get_cell_body(2, 2))

print(assigned.get_row_head(0))
print(assigned.get_cell_head(0,2))

#task = CatalogTask()
#print(str(task.get_state()))