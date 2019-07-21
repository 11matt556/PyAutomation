from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import selenium.common.exceptions as seleniumExceptions
import selenium.webdriver.support.select
import traceback

import time
import csv

# Start global variables
driver = webdriver.Chrome()
saveTicket = False
driver.implicitly_wait(3)
reviewRequired = []
verboseLog = True
actionType = None

# End global variables

# Start Functions

# === GETTERS === #


def getTicketStateStr():
    return getTicketStateObj().first_selected_option.text


def getTicketStateObj():
    return selenium.webdriver.support.select.Select(driver.find_element_by_id('sc_task.state'))


def getTicketAssigned():
    return str(driver.find_element_by_id('sys_display.sc_task.assigned_to').get_attribute("value"))


def getRepairTypeStr():
    return str(getRepairTypeObj().first_selected_option.text)


def getRepairTypeObj():
    return selenium.webdriver.support.select.Select(driver.find_element_by_xpath("//div[2]/table/tbody/tr/td/div/div/div/div[2]/select"))


def getTicketRITMStr():
    return getTicketRITMObj().get_attribute("value")


def getTicketRITMObj():
    return driver.find_element_by_id('sys_display.sc_task.request_item')


def getTicketTaskNameObj():
    return driver.find_element_by_id('sys_readonly.sc_task.u_task_name')


def getTicketTaskNameStr():
    return getTicketTaskNameObj().get_attribute("value")


def getTicketConfigurationItemObj():
    return driver.find_element_by_xpath("//*[@id='sys_display.sc_task.cmdb_ci']")


def getTicketConfigurationItemStr():
    return getTicketConfigurationItemObj().get_attribute("value")


# === SETTERS === #


# Select the type of task
# Valid tasks are Restock, Repair or Decommission
def setVariableActionType(action):
    print("Setting ticket action as " + action)
    driver.find_elements_by_xpath("//div[@class='sc_variable_editor']")[1].find_element_by_class_name(
        "input_controls").find_element_by_xpath("//option[. = '" + action + "']").click()


# Set the status of the ticket
# Valid status are 'Open' , 'Work in Progress' , 'Closed Complete' , 'Closed Incomplete' , and 'Pending'
def setTicketState(status):
    print("Setting ticket state to " + status)
    driver.find_element_by_id('sc_task.state').find_element_by_xpath("//option[. = '" + status + "']").click()


# Set the repair type for this ticket
# Valid repair types are 'onsite' and 'decommission'
def setVariableRepairType(str):
    selenium.webdriver.support.select.Select(driver.find_element_by_xpath("//div[2]/table/tbody/tr/td/div/div/div/div[2]/select")).select_by_value(str)


# Set the repair location (IE, Can this item be repaired at ISC?)
# Valid options are 'Yes' (This item can be repaired at ISC) or 'No' (This item must be repaired at MDC)
def setVariableRepairAtISC(str):
    selenium.webdriver.support.select.Select(driver.find_element_by_xpath("//tr[3]/td/div/div/div/div[2]/select")).select_by_value(str)


# Enter commentString into the Additional Comments box
def setComment(commentString):
    print("Typing " + str(commentString))
    commentArea = driver.find_element_by_xpath("//textarea[@id='activity-stream-comments-textarea']")
    # Sometimes the comment box area is collapsed by default, so we must click the button to expand it first
    try:
        commentArea.send_keys(commentString)
    except seleniumExceptions.ElementNotInteractableException:
        driver.find_element_by_xpath("//span[2]/span/nav/div/div[2]/span/button").click()
        commentArea.send_keys(commentString)


# Save changes to the ticket
def submitTicket():
    actions = webdriver.ActionChains(driver)

    saveMenu = driver.find_element_by_xpath("//nav/div")
    actions.move_to_element(saveMenu)
    actions.context_click(saveMenu)
    actions.perform()
    actions.reset_actions()

    saveButton = driver.find_element_by_xpath("//div[@id='context_1']/div[2]")

    if saveTicket:
        saveButton.click()
        print("Ticket saved.")
    else:
        print("Ticket not saved! Make sure to set the saveTicket variable to True if you want to actually submit changes to the ticket.")

# === CSV HELPER FUNCTIONS === #

def importCSV(inputCSV):
    print("Importing items...")
    with open(inputCSV) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        outputArray = []
        for row in csv_reader:
            #serial = row[0]
            #if row[1]:
            #    host = row[1]
            #if row[2]:
            #    ritm = row[2]
            outputArray.append(row[0])
            print("Importing line: " + str(row))
            line_count += 1
        print(f'Imported {line_count} items')
        print("Result: " + str(outputArray))
        return outputArray


def appendToCSV(rowArray, csvFile):
    with open(csvFile, 'a', newline='') as writeFile:
        writer = csv.writer(writeFile)
        writer.writerow(rowArray)
    writeFile.close()

# === OTHER HELPER FUNCTIONS === #

def clearCSV(csvFile):
    writeFile = open(csvFile, "w+")
    writeFile.close()


def search(inputString):
    print("Searching for " + inputString)
    searchbox = driver.find_element_by_xpath("//*[@id='sysparm_search']")
    searchbox.clear()
    searchbox.click()
    searchbox.send_keys(inputString)
    searchbox.send_keys(Keys.RETURN)


def switchToMainFrame():
    driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))


def switchToDefaultFrame():
    driver.switch_to.default_content()


def waitForAlert():
    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present(), 'Timed out waiting for confirmation popup to appear.')
        alertObj = driver.switch_to.alert
        alertObj.accept()
        print("Accepted Alert")
    except seleniumExceptions.TimeoutException as e:
        #print("No alert to accept")
        pass


def waitForTaskSelection():
    try:
        print("Waiting for task selection")
        element = WebDriverWait(driver, 10000).until(
            EC.presence_of_element_located((By.ID, "activity-stream-comments-textarea"))
            )
    except:
        print("No task selected. Quitting due to timeout")


def tableToArray(table_id):
    print("Loading table")
    table = driver.find_element_by_id(table_id)
    table_head = table.find_element_by_tag_name("thead")
    table_body = table.find_element_by_tag_name("tbody")

    elements = []
    rows = table_body.find_elements_by_tag_name("tr")
    head_rows = table_head.find_elements_by_tag_name("tr")
    head_cols = head_rows[0].find_elements_by_tag_name("th")

    rowLen = len(rows)
    colLen = len(rows[0].find_elements_by_tag_name('td')) - 1

    #print(str(colLen))

    elements.append([])
    for c in range(colLen):
        #print(str(0) + "," + str(c))
        col = head_rows[0].find_elements_by_tag_name("th")
        cell = col[c]
        elements[0].append(cell)
        #print(cell.text + " " + str(0) + "," + str(c))

    for r in range(rowLen):
        elements.append([])
        for c in range(colLen):
            col = rows[r].find_elements_by_tag_name("td")
            cell = col[c]
            elements[r+1].append(cell)
            #print(cell.text)

    return elements


def findColIndices(table,colNames):
    colDict = {}

    for c in range(len(table[0])):
        tableItem = str(table[0][c].text)

        for n in range(len(colNames)):
            if tableItem.startswith(colNames[n]):
                colDict[colNames[n]] = c

    return colDict


# List should be in the format of [['key', 'value', bool],[...]]
def clickTableItem(table_id, listOfStuff, clickThis):

    columnNames = []
    for item in listOfStuff:
        columnNames.append(item[0])
    table = tableToArray(table_id)
    columnIndices = findColIndices(table, columnNames+[clickThis])
    print("Looking through table...")
    for row in table:
        for key, value, equal in listOfStuff:
            print(row[columnIndices[key]].text)
            if equal:
                if row[columnIndices[key]].text == value:
                    continue
                else:
                    break
            else:
                if row[columnIndices[key]].text != value:
                    continue
                else:
                    break
        else:
            print("Found ticket!")
            row[columnIndices[clickThis]].click()
            break


# === MAIN TASK FUNCTIONS === #


def singleStage(taskName, action, item):
    if taskName == "rhs_restock":
        switchToDefaultFrame()
        search(item)
        switchToMainFrame()
        #  Click on the catalog tasks page
        driver.find_element_by_xpath("//div[7]/div/div/table/tbody/tr/td[2]/span/strong/a").click()

        # Locate the ticket on the catalog page
        conditions = [["State", "Open", True], ["Assignment Group", "Device Configuration (Epic)", False]]
        clickTableItem("sc_task_table", conditions, "Request item")  # Select RITM Automatically
        clickTableItem("sc_req_item.sc_task.request_item_table", conditions, "Number")  # Select the Task automatically

        # Make sure the ticket is Open
        if getTicketStateStr() != 'Open':
            raise Exception

        # Make sure this is a restock ticket
        if getTicketTaskNameStr() != taskName:
            raise Exception

        # Mark ticket as work in progress
        setTicketState('Work in Progress')
        setComment("Acknowledging asset/ticket")
        submitTicket()
        switchToDefaultFrame()
        switchToMainFrame()

        if action == "Restock":
            # Close Ticket
            setTicketState('Closed Complete')
            setComment("Device to be restocked, SSO Imprivata Project G2-G4 Minis")
            setVariableActionType('Restock')  # Set ticket task (Restock, Decommission, or Repair)
        if action == "Decom":
            # Close Ticket
            setTicketState('Closed Complete')
            setComment("This device is being decomissioned, won't be used for SSO, or is out of commission; send to MDC ")
            setVariableActionType('Decommission')  # Set ticket task (Restock, Decommission, or Repair)
        if action == "Repair (MDC)":
            # Close ticket
            setTicketState('Closed Complete')
            setComment('Laptop processed for repair / SSD swap at MDC.')
            setVariableActionType('Repair')
            setVariableRepairAtISC('No')
        if action == "Repair (ISC)":
            # Close ticket
            setTicketState('Closed Complete')
            setComment('Device to be repaired with SSD swapout; SSO Imprivata Project USDT Devices')
            setVariableActionType('Repair')
            setVariableRepairAtISC('Yes')


def outputCSV(labelType):
    appendToCSV(['', '', getTicketConfigurationItemStr(), getTicketRITMStr(), labelType], 'output.csv')
    print("RITM for " + getTicketConfigurationItemStr() + " is " + getTicketRITMStr())


def multiStage(finalTaskName, action, item):
    singleStage(finalTaskName, action, item)
    conditions = [["State", "Open", True], ["Assignment Group", "Device Configuration (Epic)", False]]
    clickTableItem("sc_req_item.sc_task.request_item_table", conditions, "Number")  # Select the Task automatically

    if finalTaskName == "rhs_restock":
        if getTicketTaskNameStr() != finalTaskName:
            raise Exception
        if action == "Restock (ISC)":
            setTicketState('Closed Complete')
            setComment("Device to be repaired with SSD swapout; SSO Imprivata Project USDT Devices")
            setVariableRepairType('onsite')


def decomItem(item):

    singleStage("rhs_restock", "Decom",item)
    outputCSV("Decommission")
    submitTicket()


def restockItem(item):

    singleStage("rhs_restock", "Restock", item)
    outputCSV("Restock")
    submitTicket()


def repairItemISC(item):

    multiStage("rhs_repair", "Repair (ISC)", item)
    # Save RITM for current item to CSV
    outputCSV("Restock")
    submitTicket()


def repairItemMDC(item):

    singleStage("rhs_restock","Repair (MDC)", item)
    # Save RITM for current item to CSV
    outputCSV("Restock")
    submitTicket()

#
#
# End Functions
#
#


# Clear any previous runs from the csv
clearCSV('output.csv')
# Set up CSV as Tim's script expects it
appendToCSV(['Tech Name', 'SN', 'PC Name', 'RITM', 'Restock/Repair', 'Label Notes'], 'output.csv')
# Import list of computer hostnames from csv
hostnames = importCSV('input.csv')

# Prompt user for type of action to take
actionType = input("\nWhat should I do with these items?\n1 = Restock, 2 = Repair / SSD Swap (ISC), 3 = Laptop Repair (MDC), 4 = Decommission\n")


# Iterate through each item and do the things
for i in range(len(hostnames)):
    print("-==========- " + hostnames[i] + " -==========-")
    print(str(i+1) + " of " + str(len(hostnames)))

    # Open to service catalog. This is to ensure each loop begins at a known starting position.
    driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Fcatalog_home.do%3Fsysparm_view%3Dcatalog_default")

    waitForAlert()
    try:
        if actionType == '1':
            restockItem(hostnames[i])
        elif actionType == '2':
            repairItemISC(hostnames[i])
        elif actionType == '3':
            print("MDC Repair not yet implemented")
            driver.close()
            exit()
        elif actionType == '4':
            decomItem(hostnames[i])
            exit()
        else:
            print("Invalid option")
            driver.close()
            exit()
    except Exception as bad:
        print("Error in item: " + hostnames[i])
        print("\t"+repr(bad))
        reviewRequired.append(hostnames[i])
        if verboseLog:
            traceback.print_tb(bad.__traceback__)
    finally:
        if reviewRequired:
            clearCSV('review.csv')
            print("THE FOLLOWING ITEMS REQUIRE MANUAL REVIEW! A copy has been saved to review.csv")
            for x in range(len(reviewRequired)):
                print("\t"+reviewRequired[x])
                appendToCSV([reviewRequired[x]], 'review.csv')