from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import selenium.common.exceptions as seleniumExceptions
import selenium.webdriver.support.select as select
import traceback
import shutil
import datetime
import os
import argparse
import time
import csv

driver = webdriver.Chrome()

def switchToContentFrame():
    driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))


def switchToDefaultFrame():
    driver.switch_to.default_content()


class Notes:

    def setAdditionalComments(self, commentString):
        switchToDefaultFrame()
        switchToContentFrame()
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

    def get_task_name(self):
        return self.taskName

class CatalogTask:
    notesTab = None
    detailsTab = None
    variablesTab = None

    def __init__(self):
        switchToDefaultFrame()
        switchToContentFrame()
        self.notesTab = Notes()
        self.detailsTab = Details()
        self.number = driver.find_element_by_id("sc_task.number").get_attribute("value")
        self.requestItem = driver.find_element_by_id("sys_display.sc_task.request_item").get_attribute("value")
        self.shortDescription = driver.find_element_by_id("sc_task.short_description").get_attribute("value")
        self.state = select.Select(driver.find_element_by_id('sc_task.state'))
        self.assignedTo = driver.find_element_by_id('sys_display.sc_task.assigned_to').get_attribute("value")

    def get_state(self):
        return self.state

    def set_state(self, state):
        print("Setting ticket state to " + state)
        driver.find_element_by_id('sc_task.state').find_element_by_xpath("//option[. = '" + state + "']").click()
        self.state = state

#For Catalog Tasks of type dm_restock
class DMRestockVariables:
    def __init__(self):
        self.restock_decom_repair = select.Select(driver.find_element_by_xpath("//tr[5]/td/div/div/div/div[2]/select"))

    def set_restock_decom_repair(self, value):
        try:
            self.restock_decom_repair.select_by_value(value)
        except seleniumExceptions.ElementNotInteractableException:
            driver.find_element_by_xpath("//span[3]/span/nav/div/div[2]/span/button").click()
            self.restock_decom_repair.select_by_value(value)


#For Catalog Tasks of type dm_repair
class DMRepairVariables:
    pass


class RestockTask(CatalogTask):
    def __init__(self):
        CatalogTask.__init__(self)
        self.variablesTab = DMRestockVariables()

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
        switchToContentFrame()
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
        #print("Getting row " + str(r))
        #for item in row_items:
            #print(item.text)
        return row_items

    def get_row_head(self, r):
        row_items = self.rows_head[r].find_elements_by_tag_name("th")
        #print("Getting row " + str(r))
        #for item in row_items:
            #print(item.get_attribute("glide_label"))
        return row_items

    def get_cell_body(self, r, c):
        cell = self.get_row_body(r)[c]
        #print("row:" + str(r) + " col:" + str(c) + " is " + cell.text)
        return cell

    def get_cell_head(self, r, c):
        cell = self.get_row_head(r)[c]
        #print("row:" + str(r) + " col:" + str(c) + " is " + cell.get_attribute("glide_label"))
        return cell

    def get_col_name(self, c) -> str:
        return str(self.get_cell_head(0, c).get_attribute("glide_label"))

    def get_col_by_name(self, name):
        for i in range(self.get_col_len_head()):
            if self.get_col_name(i).lower() == name.lower():
                return i

    def get_cell_by_column_and_name(self, colName, cellName):
        for row in range(self.get_row_len_body()):
            col = self.get_col_by_name(colName)
            cell = self.get_cell_body(row, col)
            print(str(row) + "," +str(col) + " " + cell.text)

            if str(cell.text).lower() == cellName.lower():
                print("FOUND CELL MATCHING CRITERIA [column=" + colName + " text=" + cellName + "]")
                print(cell)
                return cell


ServiceNow.homepage()

#Login
driver.find_element_by_xpath("//*[@id='maincontent']/tbody/tr[4]/td[2]").click()

#Go to a task
#driver.get("https://prismahealth.service-now.com/nav_to.do?uri=task.do?sys_id=b537bb9e1bbff308f15a4005bd4bcb8b")

#Go to tickets assigned to Tim"

ServiceNow.tickets_assigned_to_tim()

tims_table = Table("task_table")

#Find cell with task SCTASK0031734
tims_table.get_cell_by_column_and_name("Number","SCTASK0031734").click()

driver.get("https://prismahealth.service-now.com/nav_to.do?uri=%2Fsc_task.do%3Fsys_id%3Df59190fc1bc48410f15a4005bd4bcb9b%26sysparm_view%3Dtext_search")

dm_restock = RestockTask()

dm_restock.notesTab.setAdditionalComments("This is a test")

#task = CatalogTask()
#print(str(task.get_state()))