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

__CANNED_RESPONSES = {
    "acknowledge":      "Acknowledging asset/ticket",
    "restock":          "Device to be restocked, W10 Rollout; Device has been tested and wiped",
    "reclaim":          "Device has been reclaimed; Device to be restocked after testing & cleaning.",
    "decommission":     "This device is being decommissioned; send to MDC",
    "repair_mdc":       "Device to be sent in for repair/stock at MDC",
    "repair_isc_ssd":   "USDT device to be repaired with SSD swapout"
}

__VALID_STATES = {
    "wip": "Work in Progress",
    "cc": "Closed Complete"
}

SAVE_TICKET = True

driver = webdriver.Chrome()

driver.implicitly_wait(1)


def switchToContentFrame():
    driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))


def switchToDefaultFrame():
    driver.switch_to.default_content()


class CSV:
    @staticmethod
    def import_csv(inputCSV):
        print("Importing items...")
        with open(inputCSV) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            outputArray = []
            for row in csv_reader:
                outputArray.append(row)
                print("Importing line: " + str(row))
                line_count += 1

            print(f'Imported {line_count} items')
            print("Result: " + str(outputArray))
            return outputArray

    @staticmethod
    def appendToCSV(row, csvFile):
        with open(csvFile, 'a', newline='') as writeFile:
            writer = csv.writer(writeFile)
            writer.writerow(row)
        writeFile.close()


    @staticmethod
    def clearCSV(csvFile):
        writeFile = open(csvFile, "w+")
        writeFile.close()

class CatalogTask:
    notes_tab = None
    details_tab = None
    variables_tab = None

    affected_ci = None
    support_ola = None
    customer_sla = None
    all_attachments = None
    additional_catalog_tasks = None
    mce_students = None
    external_users = None

    def __init__(self):
        switchToDefaultFrame()
        switchToContentFrame()
        self.notes_tab = Notes()
        self.details_tab = Details()
        self.number = driver.find_element_by_id("sc_task.number").get_attribute("value")
        self.requestItem = driver.find_element_by_id("sys_display.sc_task.request_item").get_attribute("value")
        self.shortDescription = driver.find_element_by_id("sc_task.short_description").get_attribute("value")
        self.state = select.Select(driver.find_element_by_id('sc_task.state'))
        self.assignedTo = driver.find_element_by_id('sys_display.sc_task.assigned_to').get_attribute("value")
        self.additional_catalog_tasks = Table()
    def get_state(self):
        return self.state

    def get_ritm(self) -> str:
        return self.requestItem

    def set_state(self, state):
        print("Setting ticket state to " + state)
        driver.find_element_by_id('sc_task.state').find_element_by_xpath("//option[. = '" + state + "']").click()
        self.state = state

    def submit(self):
        actions = webdriver.ActionChains(driver)

        saveMenu = driver.find_element_by_xpath("//nav/div")
        actions.move_to_element(saveMenu)
        actions.context_click(saveMenu)
        actions.perform()
        actions.reset_actions()

        saveButton = driver.find_element_by_xpath("//div[@id='context_1']/div[2]")

        if SAVE_TICKET:
            saveButton.click()
            print("Ticket Submitted")
        else:
            print(
                "Ticket not saved! Make sure to set the saveTicket variable to True if you want to actually submit changes to the ticket.")


class Details:
    actual_start_button = None
    sequence = None
    serviceType = None
    taskName = None
    actualEnd = None
    location = None
    locationNotes = None

    def __init__(self):
        self.taskName = driver.find_element_by_id("sc_task.u_task_name").get_attribute("value")
        self.actual_start_button = driver.find_element_by_xpath("//div[2]/div/div/div/div[2]/div[2]/span/span/button")

    def get_task_name(self):
        return self.taskName

    def get_actual_start_button(self):
        return self.actual_start_button

    def accept_actual_start(self):
        driver.find_element_by_xpath("//button[@id='GwtDateTimePicker_ok']").click()

class Notes:

    def setAdditionalComments(self, commentString):
        print("Typing " + str(commentString))
        commentArea = driver.find_element_by_xpath("//textarea[@id='activity-stream-comments-textarea']")
        # Sometimes the comment box area is collapsed by default, so we must click the button to expand it first
        try:
            commentArea.send_keys(commentString)
        except seleniumExceptions.ElementNotInteractableException:
            driver.find_element_by_xpath("//span[2]/span/nav/div/div[2]/span/button").click()
            commentArea.send_keys(commentString)

#For Catalog Tasks of type dm_repair
class RepairVar:
    pass

class DmRepair(CatalogTask):
    #Stuff for repair
    pass


#For Catalog Tasks of type dm_restock
class RestockVar:
    action_required = None
    field_tech = None
    assigned_group = None
    staging_location = None

    computer_name = None
    asset_tag = None
    restock_decom_repair = None
    complete_at = None #ISC or MDC

    def __init__(self):
        try: #Try to get the action selection menu
            self.restock_decom_repair = select.Select(driver.find_element_by_xpath("//tr[5]/td/div/div/div/div[2]/select"))
        # If we can't get the menu, try expanding the parent menu out
        except seleniumExceptions.ElementNotInteractableException:
            driver.find_element_by_xpath("//span[3]/span/nav/div/div[2]/span/button").click()
            self.restock_decom_repair = select.Select(driver.find_element_by_xpath("//tr[5]/td/div/div/div/div[2]/select"))

    #Valid values are "Restock" "Repair" and "Decommission"
    def select_restock_decom_repair(self, value):
            self.restock_decom_repair.select_by_value(value)


    def select_complete_at(self, value):
        dropdown = select.Select(driver.find_element_by_xpath("//tr[3]/td/div/div/div/div[2]/select"))

        if value == "ISC":
           dropdown.select_by_value("Yes")
        elif value == "MDC":
            dropdown.select_by_value("No")


class DmRestock(CatalogTask):
    def __init__(self):
        CatalogTask.__init__(self)
        self.variables_tab = RestockVar()


class ServiceNow:
    @staticmethod
    def tims_queue():
        driver.get("https://prismahealth.service-now.com/nav_to.do?uri=%2Ftask_list.do%3Fsysparm_fixed_query%3D%26sysparm_query%3Dactive%253Dtrue%255Eassigned_to%253D78ae421e1b4fb700f15a4005bd4bcb8d%26sysparm_clear_stack%3Dtrue")

    @staticmethod
    def device_management():
        driver.get("https://prismahealth.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D9b8d095f4f9e7b0416d5334d0210c7b4%26sysparm_link_parent%3D7edd9ad21b9f734033eb85507e4bcb98%26sysparm_catalog%3D816d92d21b9f734033eb85507e4bcbf6%26sysparm_catalog_view%3Dcatalog_desktop_support")

    @staticmethod
    def homepage():
        driver.get("https://prismahealth.service-now.com")

    @staticmethod
    def waitForAlert():
        try:
            WebDriverWait(driver, 1).until(EC.alert_is_present(), 'Timed out waiting for confirmation popup to appear.')
            alertObj = driver.switch_to.alert
            alertObj.accept()
            print("Accepted Alert")
        except seleniumExceptions.TimeoutException as e:
            # print("No alert to accept")
            pass

    @staticmethod
    def search(inputString):
        switchToDefaultFrame()
        print("Searching for " + inputString)
        searchbox = driver.find_element_by_xpath("//*[@id='sysparm_search']")
        searchbox.clear()
        searchbox.click()
        searchbox.send_keys(inputString)
        searchbox.send_keys(Keys.RETURN)
        switchToDefaultFrame()
        switchToContentFrame()

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

    def get_header_row_len(self):
        return len(self.rows_head)

    def get_header_col_len(self):
        return len(self.rows_head[0].find_elements_by_tag_name('th')) - 1

    def get_header_row(self, r):
        row_items = self.rows_head[r].find_elements_by_tag_name("th")
        #print("Getting row " + str(r))
        #for item in row_items:
            #print(item.get_attribute("glide_label"))
        return row_items

    def get_header_cell(self, r, c):
        cell = self.get_header_row(r)[c]
        #print("row:" + str(r) + " col:" + str(c) + " is " + cell.get_attribute("glide_label"))
        return cell

    def get_body_row_len(self):
        return len(self.rows_body)

    def get_body_col_len(self):
        return len(self.rows_body[0].find_elements_by_tag_name('td')) - 1

    def get_body_row(self, r):
        row_items = self.rows_body[r].find_elements_by_tag_name("td")
        #print("Getting row " + str(r))
        #for item in row_items:
            #print(item.text)
        return row_items

    def get_body_cell(self, r, c):
        cell = self.get_body_row(r)[c]
        #print("row:" + str(r) + " col:" + str(c) + " is " + cell.text)
        return cell

    def get_col_name(self, c) -> str:
        return str(self.get_header_cell(0, c).get_attribute("glide_label"))

    def get_col_by_name(self, name) -> int:
        for i in range(self.get_header_col_len()):
            if self.get_col_name(i).lower() == name.lower():
                return i

# Returns row of cell, not the cell itself!
    def find_cell_in_col(self, colName, cellName) -> list:
        for row in range(self.get_body_row_len()):
            col = self.get_col_by_name(colName)
            cell = self.get_body_cell(row, col)
            print(str(row) + "," +str(col) + " " + cell.text)

            if str(cell.text).lower() == cellName.lower():
                print("FOUND CELL!")
                return row


def doDecom(hostname):
    print("DECOMMISSIONING " + hostname)
    dm_restock = DmRestock()
    #dm_restock.set_state(__VALID_STATES['wip'])
    #dm_restock.notes_tab.setAdditionalComments(__CANNED_RESPONSES["acknowledge"])
    #dm_restock.submit()

    dm_restock.set_state(__VALID_STATES['cc'])
    dm_restock.notes_tab.setAdditionalComments(__CANNED_RESPONSES['decommission'])
    dm_restock.variables_tab.select_restock_decom_repair("decommission")
    ritm = dm_restock.get_ritm()
    dm_restock.details_tab.get_actual_start_button().click()
    #time.sleep(3)
    dm_restock.details_tab.accept_actual_start()
    #time.sleep(3)
    dm_restock.submit()

    CSV.appendToCSV(['', '', hostname, dm_restock.get_ritm(), "Decommission"], 'output.csv')


def doRestock(hostname):
    print("RESTOCKING " + hostname)
    dm_restock = DmRestock()
    #dm_restock.set_state(__VALID_STATES['wip'])
    #dm_restock.notes_tab.setAdditionalComments(__CANNED_RESPONSES["acknowledge"])
    #dm_restock.submit()

    dm_restock.set_state(__VALID_STATES['cc'])
    dm_restock.notes_tab.setAdditionalComments(__CANNED_RESPONSES['restock'])
    dm_restock.variables_tab.select_restock_decom_repair("restock")
    ritm = dm_restock.get_ritm()
    dm_restock.details_tab.get_actual_start_button().click()
    #time.sleep(3)
    dm_restock.details_tab.accept_actual_start()
    #time.sleep(3)
    dm_restock.submit()

    CSV.appendToCSV(['', '', hostname, dm_restock.get_ritm(), "restock"], 'output.csv')

# repar_type can be "isc", or "mdc"
def doRepair(hostname, repair_type):
    print("REPAIRING " + hostname)
    dm_restock = DmRestock()

    dm_restock.set_state(__VALID_STATES['cc'])

    dm_restock.variables_tab.select_restock_decom_repair("repair")
    dm_restock.details_tab.get_actual_start_button().click()

    ritm = dm_restock.get_ritm()
    if repair_type == "isc":
        dm_restock.notes_tab.setAdditionalComments(__CANNED_RESPONSES['repair_isc_ssd'])
        dm_restock.variables_tab.select_complete_at("isc")
        #Go to RITM page
        ServiceNow.search(ritm)
        #Find and go to open repair task
        table = Table("sc_req_item.sc_task.request_item_table")
        state_col = table.get_col_by_name("State")
        task_col = table.get_col_by_name("Number")
        desc_col = table.get_col_by_name("Short description")

        cell_row = table.find_cell_in_col("state", "open")

        table.get_body_cell(cell[0],task_col).click()

        #Complete repair

    elif repair_type == "mdc":
        dm_restock.notes_tab.setAdditionalComments((__CANNED_RESPONSES['mdc']))
        dm_restock.variables_tab.select_complete_at("mdc")



    #time.sleep(3)
    dm_restock.details_tab.accept_actual_start()
    #time.sleep(3)

    raise NotImplementedError
    #dm_restock.submit()

computers = CSV.import_csv('input.csv')

ServiceNow.homepage()

#Login
driver.find_element_by_xpath("//*[@id='maincontent']/tbody/tr[4]/td[2]").click()
CSV.clearCSV('output.csv')
CSV.appendToCSV(['Tech Name', 'SN', 'PC Name', 'RITM', 'Restock/Repair', 'Label Notes'], 'output.csv')

for item in computers:
    hostname = item[0]
    task = item[1]

    ServiceNow.tims_queue()

    if SAVE_TICKET == False:
        ServiceNow.waitForAlert()

    tims_table = Table("task_table")
    item_row = tims_table.find_cell_in_col("Configuration item", hostname)
    taskCol = tims_table.get_col_by_name("number")
    tims_table.get_body_cell(item_row, taskCol).click()

    try:
        if task == "Decommission":
            doDecom(hostname)
        elif task == "Restock":
            pass
    except Exception as e:
        print(hostname+": "+ "Something went wrong!")
        traceback.print_tb(e.__traceback__)