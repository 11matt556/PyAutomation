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
saveTicket = True
driver.implicitly_wait(3)
timeToSleep = 0
reviewRequired = []
verboseLog = True
# End global variables

# Start Functions


def submitTicket():
    actions = webdriver.ActionChains(driver)
    # print("Switch to iframe")
    # driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))

    saveMenu = driver.find_element_by_xpath("//nav/div")
    # print(' Save menu is here: ' + str(saveMenu))
    actions.move_to_element(saveMenu)
    actions.context_click(saveMenu)
    actions.perform()
    actions.reset_actions()

    saveButton = driver.find_element_by_xpath("//div[@id='context_1']/div[2]")
    # print(' Save button is here: ' + str(saveButton))

    if saveTicket:
        saveButton.click()
        print("Ticked saved.")
    else:
        print("Ticket not saved! Make sure to set the saveTicket variable to True if you want to actually submit changes to the ticket.")


def importCSV(inputCSV):
    with open(inputCSV) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        outputArray = []
        for row in csv_reader:
            serial = row[0]
            #if row[1]:
            #    host = row[1]
            #if row[2]:
            #    ritm = row[2]
            outputArray.append(serial)
            print("Processing: " + str(row))
            line_count += 1
        print(f'Processed {line_count} lines:')
        print("Result: " + str(outputArray))
        return outputArray


def appendToCSV(rowArray, csvFile):
    with open(csvFile, 'a', newline='') as writeFile:
        writer = csv.writer(writeFile)
        writer.writerow(rowArray)
    writeFile.close()


def clearCSV(csvFile):
    writeFile = open(csvFile, "w+")
    writeFile.close()


# Valid tasks are Restock, Repair or Decommission
def setTaskType(task):
    print("Setting ticket task as " + task)
    driver.find_elements_by_xpath("//div[@class='sc_variable_editor']")[1].find_element_by_class_name(
        "input_controls").find_element_by_xpath("//option[. = '" + task + "']").click()


# Valid status are is 'Open' , 'Work in Progress' , 'Closed Complete' , 'Closed Incomplete' , and 'Pending'
def setTicketState(status):
    print("Setting ticket state to " + status)
    driver.find_element_by_id('sc_task.state').find_element_by_xpath("//option[. = '" + status + "']").click()


def getTicketState():
    return str(selenium.webdriver.support.select.Select(driver.find_element_by_id('sc_task.state')).first_selected_option.text)


def getTicketAssigned():
    return str(driver.find_element_by_id('sys_display.sc_task.assigned_to').get_attribute("value"))


def inputComment(commentString):
    print("Typing " + str(commentString))
    commentArea = driver.find_element_by_xpath("//textarea[@id='activity-stream-comments-textarea']")
    # commentArea.click()
    commentArea.send_keys(commentString)


def search(inputString):
    print("Searching for " + inputString)
    searchbox = driver.find_element_by_xpath("//*[@id='sysparm_search']")
    searchbox.clear()
    searchbox.click()
    searchbox.send_keys(inputString)
    searchbox.send_keys(Keys.RETURN)


def getRITM():
    return driver.find_element_by_id('sys_display.sc_task.request_item').get_attribute("value")


def switchToMainFrame():
    driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))


def switchToDefaultFrame():
    driver.switch_to.default_content()


def restockItem(item):

    # Make sure we have the default/top iframe selected
    driver.switch_to.default_content()

    # Search for the 1st computer
    search(item)

    # Switch to main content iframe
    switchToMainFrame()

    # Wait until a task is selected
    try:
        print("Waiting for task selection")
        element = WebDriverWait(driver, 10000).until(
            EC.presence_of_element_located((By.ID, "activity-stream-comments-textarea"))
            )
    except:
        print("No task selected. Quitting due to timeout")

    time.sleep(timeToSleep)

    # Make sure the ticket is Open
    if getTicketState() != 'Open':
        raise Exception

    # Make sure the ticket is assigned to Bryan or John
    if getTicketAssigned() != ('Bryan Shain' or 'John Higman'):
        raise Exception

    # Set ticket state
    setTicketState('Work in Progress')
    time.sleep(timeToSleep)

    # Enter comment
    # Sometimes the comment box area is collapsed by default, so we must click the button to expand it first
    try:
        inputComment("Acknowledging asset/ticket")
        time.sleep(timeToSleep)
    except seleniumExceptions.ElementNotInteractableException:
        driver.find_element_by_xpath("//span[2]/span/nav/div/div[2]/span/button").click()

    # Save ticket
    submitTicket()
    time.sleep(timeToSleep)

    driver.switch_to.default_content()

    # Switch to main content iframe
    driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))

    # Set ticket state
    setTicketState('Closed Complete')
    time.sleep(timeToSleep)

    # Enter comment
    inputComment("Device to be restocked, SSO Imprivata Project G2-G4 Minis")
    time.sleep(timeToSleep)

    # Set ticket task (Restock, Decommission, or Repair)
    setTaskType('Restock')
    time.sleep(timeToSleep)

    # Save RITM for current item to CSV
    appendToCSV([item,getRITM()], 'output.csv')
    print("RITM for " + item + " is " + getRITM())

    time.sleep(timeToSleep)

    # Save ticket
    submitTicket()

# Returns a string, not the actual object
def getRepairTypeStr():
    return str(getRepairTypeObj().first_selected_option.text)

def getRepairTypeObj():
    return selenium.webdriver.support.select.Select(driver.find_element_by_xpath("//div[2]/table/tbody/tr/td/div/div/div/div[2]/select"))

def setRepairType(str):
    selenium.webdriver.support.select.Select(driver.find_element_by_xpath("//div[2]/table/tbody/tr/td/div/div/div/div[2]/select")).select_by_value(str)

def setRepairLocation(str):
    selenium.webdriver.support.select.Select(driver.find_element_by_xpath("//tr[3]/td/div/div/div/div[2]/select")).select_by_value(str)

def waitForTaskSelection():
    try:
        print("Waiting for task selection")
        element = WebDriverWait(driver, 10000).until(
            EC.presence_of_element_located((By.ID, "activity-stream-comments-textarea"))
            )
    except:
        print("No task selected. Quitting due to timeout")

def repairItem(item):
    #driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Fsc_task.do%3Fsys_id%3D86f33381db6c93c8cf1fa961ca9619c1%26sysparm_view%3Dtext_search%26sysparm_record_target%3Dsc_task%26sysparm_record_row%3D1%26sysparm_record_rows%3D4362%26sysparm_record_list%3D123TEXTQUERY321%25253Drhs_repair")

    # Make sure we have the default/top iframe selected
    driver.switch_to.default_content()

    # Search for the computer
    search(item)

    # Switch to main content iframe
    switchToMainFrame()

    # Wait until a task is selected
    waitForTaskSelection()
    time.sleep(timeToSleep)

    # Make sure the ticket is Open
    if getTicketState() != 'Open':
        raise Exception

    # Make sure the ticket is assigned to Bryan or John
    if getTicketAssigned() != ('Bryan Shain' or 'John Higman'):
        raise Exception

    setTicketState('Work in Progress')

    inputComment('Acknowledging asset/ticket')

    submitTicket()

    setTicketState('Closed Complete')

    inputComment('Device to be repaired with SSD swapout; SSO Imprivata Project USDT Devices')

    setTaskType('Repair')

    setRepairLocation('Yes')

    submitTicket()

    ritm = getRITM()

    switchToDefaultFrame()

    search(ritm)

    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present(), 'Timed out waiting for confirmation popup to appear.')
        alert = driver.switch_to.alert
        alert.accept()
        print("Accepted Alert")
    except seleniumExceptions.TimeoutException as e:
        print("No alert to accept")

    switchToMainFrame()

    driver.find_element(By.LINK_TEXT, "Click here").click()
    #driver.find_element_by_xpath("//div[@id='output_messages']/div/div/div/a").click()
    # Wait until a task is selected
    try:
        print("Waiting for task selection")
        element = WebDriverWait(driver, 10000).until(
            EC.presence_of_element_located((By.ID, "activity-stream-comments-textarea"))
            )
    except:
        print("No task selected. Quitting due to timeout")

    inputComment("Device to be repaired with SSD swapout; SSO Imprivata Project USDT Devices")

    #print(driver.find_elements_by_xpath("//div[@class='sc_variable_editor']")[1].find_element_by_class_name("input_controls"))
    #print(driver.find_element_by_xpath(("//div[2]/table/tbody/tr/td/div/div/div/div[2]/select")))

    setTicketState('Closed Complete')

    # print(getRepairTypeStr())

    # onsite corresponds to "Repair Completed" as opposed to a decom repair
    setRepairType('onsite')

    # Save RITM for current item to CSV
    appendToCSV([item,getRITM()], 'output.csv')
    print("RITM for " + item + " is " + getRITM())

    submitTicket()
#
#
# End Functions
#
#

# Clear any previous runs from the csv
clearCSV('output.csv')

# Import list of computer hostnames
computers = importCSV('input.csv')

# Print out list of computer hostnames
for i in range(len(computers)):
    print("-==========- " + computers[i] + " -==========-")
    print(str(i+1) + " of " + str(len(computers)))
    # Open to service catalog
    driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Fcatalog_home.do%3Fsysparm_view%3Dcatalog_default")

    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present(), 'Timed out waiting for confirmation popup to appear.')
        alert = driver.switch_to.alert
        alert.accept()
        print("Accepted Alert")
    except seleniumExceptions.TimeoutException as e:
        print("No alert to accept")

    try:
        #restockItem(computers[i])
        repairItem(computers[i])
    except Exception as bad:
        print("Error in item: " + computers[i])
        print("\t"+repr(bad))
        reviewRequired.append(computers[i])
        if verboseLog:
            traceback.print_tb(bad.__traceback__)

if reviewRequired:
    clearCSV('review.csv')
    print("THE FOLLOWING ITEMS REQUIRE MANUAL REVIEW! A copy has been saved to review.csv")
    for x in range(len(reviewRequired)):
        print("\t"+reviewRequired[x])
        appendToCSV([reviewRequired[x]], 'review.csv')
