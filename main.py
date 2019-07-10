from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import selenium.common.exceptions as seleniumExceptions

import time
import csv

# Start global variables
driver = webdriver.Chrome()
saveTicket = False
driver.implicitly_wait(3)
timeToSleep = 0
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
    with open(csvFile, 'a') as writeFile:
        writer = csv.writer(writeFile)
        writer.writerow(rowArray)
    writeFile.close()


def clearCSV(csvFile):
    writeFile = open(csvFile, "w+")
    writeFile.close()

# Valid tasks are Restock, Repair or Decommission
def selTaskTypeFromDropdown(task):
    print("Setting ticket task as " + task)
    driver.find_elements_by_xpath("//div[@class='sc_variable_editor']")[1].find_element_by_class_name(
        "input_controls").find_element_by_xpath("//option[. = '" + task + "']").click()


# Valid status are is 'Open' , 'Work in Progress' , 'Closed Complete' , 'Closed Incomplete' , and 'Pending'
def selTicketStateFromDropdown(status):
    print("Setting ticket state to " + status)
    driver.find_element_by_id('sc_task.state').find_element_by_xpath("//option[. = '" + status + "']").click()


def inputComment(commentString):
    print("Typing " + str(commentString))
    commentArea = driver.find_element_by_xpath("//textarea[@id='activity-stream-comments-textarea']")
    # commentArea.click()
    commentArea.send_keys(commentString)


def search(inputString):
    print("Searching for " + inputString)
    searchbox = driver.find_element_by_xpath("//*[@id='sysparm_search']")
    searchbox.click()
    searchbox.send_keys(inputString)
    searchbox.send_keys(Keys.RETURN)


def saveRITMToCSV(outputCSV):
    RITM = []
    RITM.append(driver.find_element_by_id('sys_display.sc_task.request_item').get_attribute("value"))
    print("Saving " + str(RITM) + " to " + outputCSV)
    appendToCSV(RITM, outputCSV)


def restockItem(item):

    # Make sure we have the default/top iframe selected
    driver.switch_to.default_content()

    # Search for the 1st computer
    search(item)

    # Switch to main content iframe
    driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))

    # Wait until a task is selected
    try:
        print("Waiting for task selection")
        element = WebDriverWait(driver, 10000).until(
            EC.presence_of_element_located((By.ID, "activity-stream-comments-textarea"))
        )
    except:
        print("No task selected. Quitting due to timeout")

    time.sleep(timeToSleep)

    # Set ticket state
    selTicketStateFromDropdown('Work in Progress')

    time.sleep(timeToSleep)

    # Enter comment
    inputComment("Acknowledging asset/ticket")
    time.sleep(timeToSleep)

    # Save ticket
    submitTicket()
    time.sleep(timeToSleep)

    driver.switch_to.default_content()
    # Switch to main content iframe
    driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))

    # Set ticket state
    selTicketStateFromDropdown('Closed Complete')
    time.sleep(timeToSleep)

    #    Enter comment
    inputComment("Device to be restocked, SSO Imprivata Project G2-G4 Minis")
    time.sleep(timeToSleep)

    # Set ticket task (Restock, Decommission, or Repair)
    selTaskTypeFromDropdown('Restock')
    time.sleep(timeToSleep)

    # Save RITM to CSV
    saveRITMToCSV('output.csv')
    time.sleep(timeToSleep)

    # Save ticket
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
    print("Currently working on item: " + computers[i])

    # Open to service catalog
    driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Fcatalog_home.do%3Fsysparm_view%3Dcatalog_default")

    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present(), 'Timed out waiting for confirmation popup to appear.')
        alert = driver.switch_to.alert
        alert.accept()
        print("Accepted Alert")
    except seleniumExceptions.TimeoutException as e:
        print("No alert to accept")

    try:
        restockItem(computers[i])
    except Exception as bad:
        print("Something went wrong!")
        print(bad)