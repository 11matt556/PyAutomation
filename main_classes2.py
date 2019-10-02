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
    "acknowledge": "Acknowledging asset/ticket",
    "restock": "Device to be restocked, W10 Rollout; Device has been tested and wiped",
    "reclaim": "Device has been reclaimed; Device to be restocked after testing & cleaning.",
    "decommission": "This device is being decommissioned; send to MDC",
    "repair_mdc": "Device to be sent in for repair/stock at MDC",
    "repair_isc_ssd": "USDT device to be repaired with SSD swapout"
}

__VALID_STATES = {
    "wip": "Work in Progress",
    "cc": "Closed Complete"
}

REVIEW_REQUIRED = []
VERBOSE_LOG = True
SAVE_TICKET = False

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

    def get_state(self):
        return self.state

    def get_ritm(self) -> str:
        return self.requestItem

    def set_state(self, state):
        print("Setting ticket state to " + state)
        driver.find_element_by_id('sc_task.state').find_element_by_xpath("//option[. = '" + state + "']").click()
        self.state = state

    @staticmethod
    def submit():
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

    @staticmethod
    def accept_actual_start():
        driver.find_element_by_xpath("//button[@id='GwtDateTimePicker_ok']").click()

    def set_actual_start(self):
        self.get_actual_start_button().click()
        time.sleep(1)
        self.accept_actual_start()


class Notes:

    @staticmethod
    def setAdditionalComments(commentString):
        print("Typing " + str(commentString))
        commentArea = driver.find_element_by_xpath("//textarea[@id='activity-stream-comments-textarea']")
        # Sometimes the comment box area is collapsed by default, so we must click the button to expand it first
        try:
            commentArea.send_keys(commentString)
        except seleniumExceptions.ElementNotInteractableException:
            driver.find_element_by_xpath("//span[2]/span/nav/div/div[2]/span/button").click()
            commentArea.send_keys(commentString)


class RepairVar:  # Variable tab for dm_repair tickets
    def __init__(self):
        try:
            self.repair_type = select.Select(
                driver.find_element_by_xpath("//div[2]/table/tbody/tr[2]/td/div/div/div/div[2]/select"))
        except seleniumExceptions.ElementNotInteractableException:
            driver.find_element_by_xpath("//span[3]/span/nav/div/div[2]/span/button").click()
            self.repair_type = select.Select(
                driver.find_element_by_xpath("//div[2]/table/tbody/tr[2]/td/div/div/div/div[2]/select"))

    def select_repair_type(self, repair_type: str):
        if repair_type.lower() == 'restock':
            self.repair_type.select_by_value('restock')
        else:
            print(repair_type + " is not a valid option")


class DmRepair(CatalogTask):
    def __init__(self):
        CatalogTask.__init__(self)
        self.variables_tab = RepairVar()


class DmRestock(CatalogTask):
    def __init__(self):
        CatalogTask.__init__(self)
        self.variables_tab = RestockVar()


class RestockVar:  # Variable tab for dm_restock tickets
    action_required = None
    field_tech = None
    assigned_group = None
    staging_location = None

    computer_name = None
    asset_tag = None
    restock_decom_repair = None

    def __init__(self):
        try:
            self.restock_decom_repair = select.Select(
                driver.find_element_by_xpath("//tr[5]/td/div/div/div/div[2]/select"))
        # If we can't get the menu, try expanding the parent menu out
        except seleniumExceptions.ElementNotInteractableException:
            driver.find_element_by_xpath("//span[3]/span/nav/div/div[2]/span/button").click()
            self.restock_decom_repair = select.Select(
                driver.find_element_by_xpath("//tr[5]/td/div/div/div/div[2]/select"))

    def select_restock_decom_repair(self, value):  # Valid values are "Restock" "Repair" and "Decommission"
        print("Selecting " + value)
        self.restock_decom_repair.select_by_value(value)

    @staticmethod
    def select_complete_at(value):
        dropdown = select.Select(driver.find_element_by_xpath("//tr[6]/td/div/div/div/div[2]/select"))
        print("Repair to be completed at " + value)
        if value.lower() == "isc":
            dropdown.select_by_value("Yes")
        elif value.lower() == "mdc":
            dropdown.select_by_value("No")
        else:
            print(value + " is not a valid repair completion location")
            raise Exception


class ServiceNow:
    @staticmethod
    def tims_queue():
        driver.get(
            "https://prismahealth.service-now.com/nav_to.do?uri=%2Ftask_list.do%3Fsysparm_fixed_query%3D%26sysparm_query%3Dactive%253Dtrue%255Eassigned_to%253D78ae421e1b4fb700f15a4005bd4bcb8d%26sysparm_clear_stack%3Dtrue")

    @staticmethod
    def device_management():
        driver.get(
            "https://prismahealth.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D9b8d095f4f9e7b0416d5334d0210c7b4%26sysparm_link_parent%3D7edd9ad21b9f734033eb85507e4bcb98%26sysparm_catalog%3D816d92d21b9f734033eb85507e4bcbf6%26sysparm_catalog_view%3Dcatalog_desktop_support")

    @staticmethod
    def homepage():
        driver.get("https://prismahealth.service-now.com")

    @staticmethod
    def acceptAlert():
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
        searchbutton = driver.find_element_by_xpath("//label/span")
        searchbutton.click()
        searchbox = driver.find_element_by_xpath("//form/div/input")
        searchbox.clear()
        searchbox.click()
        searchbox.send_keys(inputString)
        searchbox.send_keys(Keys.RETURN)
        if not SAVE_TICKET:
            ServiceNow.acceptAlert()
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
        print("Loading table...")
        self.table = driver.find_element_by_id(table_id)
        self.table_head = self.table.find_element_by_tag_name("thead")
        self.table_body = self.table.find_element_by_tag_name("tbody")

        self.rows_body = self.table_body.find_elements_by_tag_name("tr")
        self.rows_head = self.table_head.find_elements_by_tag_name("tr")
        print("Table loaded")

    def get_header_row_len(self):  # Return the number of rows in the header. This should almost always be 1.
        return len(self.rows_head)

    def get_header_col_len(self):  # Return the number of columns in the header.
        return len(self.rows_head[0].find_elements_by_tag_name('th')) - 1

    def get_header_row(self, r):  # Return the row as a list of Selenium Web Element
        row_items = self.rows_head[r].find_elements_by_tag_name("th")
        return row_items

    def get_header_cell(self, r, c):  # Returns a single Selenium Web Element from the table
        cell = self.get_header_row(r)[c]
        return cell

    def get_body_row_len(self):
        return len(self.rows_body)

    def get_body_col_len(self):
        return len(self.rows_body[0].find_elements_by_tag_name('td')) - 1

    def get_body_row(self, r):
        row_items = self.rows_body[r].find_elements_by_tag_name("td")
        return row_items

    def get_body_cell(self, r, c):
        cell = self.get_body_row(r)[c]
        return cell

    def get_col_name(self, c) -> str:  # Return the name of a column when given the index
        return str(self.get_header_cell(0, c).get_attribute("glide_label"))

    def find_col_with_name(self, name) -> int:  # Return the index of a column when given the name of the column
        for i in range(self.get_header_col_len()):
            if self.get_col_name(i).lower() == name.lower():
                return i

    # Returns row index of cell, not the cell itself!
    def find_in_col(self, cellName,
                    colName) -> int:  # Search in coName for a cell named cellName and return the row index
        print("----Looking for " + cellName + " in " + colName + "----")
        for row_index in range(self.get_body_row_len()):
            col = self.find_col_with_name(colName)
            cell = self.get_body_cell(row_index, col)
            print(str(row_index) + "," + str(col) + " " + cell.text + "\r", end=" ", flush=True)

            if str(cell.text).lower() == cellName.lower():
                print("FOUND " + str(cellName) + " in row_index " + str(row_index))
                return row_index


def doDecom(configuration_item):
    print("DECOMMISSIONING " + configuration_item)
    dm_restock = DmRestock()
    ritm = dm_restock.get_ritm()
    dm_restock.set_state(__VALID_STATES['cc'])
    dm_restock.notes_tab.setAdditionalComments(__CANNED_RESPONSES['decommission'])
    dm_restock.variables_tab.select_restock_decom_repair("decommission")
    dm_restock.details_tab.set_actual_start()
    dm_restock.submit()

    CSV.appendToCSV(['', '', configuration_item, ritm, "Decommission"], 'output.csv')


def doRestock(configuration_item):
    print("RESTOCKING " + configuration_item)
    dm_restock = DmRestock()
    ritm = dm_restock.get_ritm()
    dm_restock.set_state(__VALID_STATES['cc'])
    dm_restock.notes_tab.setAdditionalComments(__CANNED_RESPONSES['restock'])
    dm_restock.variables_tab.select_restock_decom_repair("restock")
    dm_restock.details_tab.set_actual_start()
    dm_restock.submit()

    CSV.appendToCSV(['', '', configuration_item, ritm, "restock"], 'output.csv')


def doRepair(configuration_item, repair_type):  # repair_type can be "isc", or "mdc"
    print("REPAIRING " + configuration_item)
    dm_restock = DmRestock()
    ritm = dm_restock.get_ritm()  # Grab the ritm since we will need this later
    dm_restock.set_state(__VALID_STATES['cc'])  # Set ticket state to Closed Complete
    dm_restock.variables_tab.select_restock_decom_repair(
        "repair")  # Select 'Repair' from the REpair/Restock/Decom dropdown
    dm_restock.details_tab.set_actual_start()  # Set the actual start date to the current time

    if repair_type == "isc":
        dm_restock.notes_tab.setAdditionalComments(__CANNED_RESPONSES['repair_isc_ssd'])  # Set the comment
        dm_restock.variables_tab.select_complete_at("isc")  # Select that the ticket will be completed at isc
        dm_restock.submit()  # Submit the ticket

        # Now we need to locate the dm_repair task to complete the repair.

        ServiceNow.search(ritm)  # Go to the RITM page.

        table = Table("sc_req_item.sc_task.request_item_table")  # Load table of catalog tasks associated with this RITM
        state_col = table.find_col_with_name("State")
        task_col = table.find_col_with_name("Number")
        desc_col = table.find_col_with_name("Short description")

        cell_row = table.find_in_col("open", "state")  # Find open ticket
        # TODO: Perform additional validation that this is the correct ticket
        table.get_body_cell(cell_row, task_col).click()  # Click on the open ticket

        # Now we are on the dm_repair ticket page
        dm_repair = DmRepair()
        dm_repair.set_state(__VALID_STATES['cc'])
        dm_repair.details_tab.set_actual_start()
        dm_repair.notes_tab.setAdditionalComments(__CANNED_RESPONSES['repair_isc_ssd'])
        dm_repair.variables_tab.select_repair_type('restock')
        dm_repair.submit()

    elif repair_type == "mdc":
        dm_restock.notes_tab.setAdditionalComments((__CANNED_RESPONSES['mdc']))
        dm_restock.variables_tab.select_complete_at("mdc")
        dm_restock.submit()

    CSV.appendToCSV(['', '', configuration_item, ritm, "restock"], 'output.csv')

def write_review():
    CSV.clearCSV('review.csv')
    print("THE FOLLOWING ITEMS REQUIRE MANUAL REVIEW! A copy has been saved to review.csv")
    for x in range(len(REVIEW_REQUIRED)):
        print("\t"+str(REVIEW_REQUIRED[x]))
        CSV.appendToCSV(REVIEW_REQUIRED[x], 'review.csv')

class ItemNotFound(Exception):
    pass

computers = CSV.import_csv('input.csv')

# TODO: Check CSV for errors (eg, "decom" instead of "decommission"

ServiceNow.homepage()

# Login
driver.find_element_by_xpath("//*[@id='maincontent']/tbody/tr[4]/td[2]").click()
CSV.clearCSV('output.csv')
CSV.appendToCSV(['Tech Name', 'SN', 'PC Name', 'RITM', 'Restock/Repair', 'Label Notes'], 'output.csv')

current = 1
total_time = 0
for item in computers:
    start_time = time.time()

    ServiceNow.tims_queue()

    if not SAVE_TICKET:
        ServiceNow.acceptAlert()

    try:
        hostname = item[0]
        task = str(item[1]).lower()
        print("====================" + hostname + " (" + task + ")" + "====================")
        print(str(current) + " of " + str(len(computers)))
        tims_table = Table("task_table")
        # TODO: Use ItemNotFound exception to properly warn user when an item can't be found in a table, rather than using the generic Exception
        item_row = tims_table.find_in_col(hostname, "Configuration item") # Find which row the hostname is in
        taskCol = tims_table.find_col_with_name("number")  # Get the index of the task "number" column
        tims_table.get_body_cell(item_row,  tims_table.find_col_with_name("number")).click()

        if task == "decommission":
            doDecom(hostname)

        elif task == "restock":
            doRestock(hostname)

        elif task == "repair_isc":
            doRepair(hostname, "isc")

        elif task == "repair_mdc":
            doRepair(hostname, "mdc")

    except Exception as e:
        print("Error in item: " + hostname)
        print("\t"+repr(e))
        REVIEW_REQUIRED.append([hostname, traceback.format_exc()])

        if VERBOSE_LOG:
            traceback.print_tb(e.__traceback__)

        write_review()
    except ItemNotFound as e:
        print("Unable to find " + hostname + " in table.")
        print("\t"+repr(e))
        REVIEW_REQUIRED.append([hostname, traceback.format_exc()])

        if VERBOSE_LOG:
            traceback.print_tb(e.__traceback__)

    finally:
        if REVIEW_REQUIRED:
            write_review()

        elapsed_time = time.time() - start_time
        print(str(elapsed_time) + " taken to complete ticket")
        total_time = elapsed_time + total_time
        print("Total time taken: " + str(total_time))
        avg_time = total_time/current
        eta_time = avg_time * (len(computers) - current)
        print("ETA: " + str(eta_time))
        current = current+1
