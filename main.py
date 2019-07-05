from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import csv

# Start global variables
driver = webdriver.Chrome()
actions = webdriver.ActionChains(driver)

driver.implicitly_wait(3)
# End global variables
# Start Functions
def submitTicket(bool):
    # print("Switch to iframe")
    # driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))

    saveMenu = driver.find_element_by_xpath("//nav/div")
    # print(' Save menu is here: ' + str(saveMenu))
    actions.move_to_element(saveMenu)
    actions.context_click(saveMenu)
    actions.perform()

    saveButton = driver.find_element_by_xpath("//div[@id='context_1']/div[2]")
    # print(' Save button is here: ' + str(saveButton))

    if bool == 'true':
        actions.move_to_element(saveButton)
        actions.click(saveButton)
        actions.perform()
    else:
        print("Ticket not saved! Make sure to pass 'true' if you want to actually submit.")

def importCSV(inputCSV):
    outputArray = open(inputCSV, 'r').read(-1).split(',')
    return outputArray

def writeToCSV(rowArray, csvFile):
    with open(csvFile, 'w') as writeFile:
        writer = csv.writer(writeFile)
        writer.writerow(rowArray)
    writeFile.close()

# Valid tasks are Restock, Repair or Decommission
def selTaskTypeFromDropdown(task):
    driver.find_elements_by_xpath("//div[@class='sc_variable_editor']")[1].find_element_by_class_name(
        "input_controls").find_element_by_xpath("//option[. = '" + task + "']").click()

def selTicketStateFromDropdown(status):
    driver.find_element_by_id('sc_task.state').find_element_by_xpath("//option[. = '" + status + "']").click()

def inputComment(commentString):
    commentArea = driver.find_element_by_id('activity-stream-comments-textarea')
    actions.move_to_element(commentArea)
    actions.click(commentArea)
    actions.send_keys(commentString)
    actions.perform()
#
#
# End Functions
#
#

# Import list of computer hostnames
computers = importCSV('input.csv')

# Print out list of computer hostnames
for i in range(len(computers)):
    print(computers[i])

# Navigate to webpage
#driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Fsc_task.do%3Fsys_id%3D07feba5bdb1677c0d4aa710439961995%26sysparm_view%3Dtext_search%26sysparm_record_target%3Dsc_task%26sysparm_record_row%3D3%26sysparm_record_rows%3D4%26sysparm_record_list%3D123TEXTQUERY321%25253DDT2UA50516QQ")

driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Ftextsearch.do%3Fsysparm_search%3DDT2UA3351XQ6")

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
    driver.quit()

driver.switch_to.default_content()

# Click search box
driver.find_element_by_xpath("//*[@id='sysparm_search']").click()

# Switch to main content iframe
driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))


# Set ticket state
selTicketStateFromDropdown('Work in Progress')

# Enter comment
inputComment("Acknowledging device/asset")

# Save ticket
submitTicket('false')

# Set ticket state
selTicketStateFromDropdown('Closed Complete')

# Enter comment
inputComment("Device to be restocked")

# Set ticket task (Restock, Decommission, or Repair)
selTaskTypeFromDropdown('Restock')

# Save RITM to CSV
RITM = driver.find_element_by_id('sys_display.sc_task.request_item').get_attribute("value")
print(RITM)
writeToCSV(RITM, 'output.csv')

# Save ticket
submitTicket('false')
