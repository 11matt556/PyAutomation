from selenium import webdriver
import time
import csv

# Start global variables
driver = webdriver.Chrome()
actions = webdriver.ActionChains(driver)
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
        print("Ticket not submitted! Make sure to pass 'true' if you want to actually submit.")

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

def selTicketStatusFromDropdown(status):
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
computers = importCSV('input.csv')

for i in range(len(computers)):
    print(computers[i])

driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Fsc_task.do%3Fsys_id%3D07feba5bdb1677c0d4aa710439961995%26sysparm_view%3Dtext_search%26sysparm_record_target%3Dsc_task%26sysparm_record_row%3D3%26sysparm_record_rows%3D4%26sysparm_record_list%3D123TEXTQUERY321%25253DDT2UA50516QQ")
print("Paused")
time.sleep(10)
print("Unpaused")

driver.find_element_by_xpath("//*[@id='sysparm_search']").click()

# Switch to main content iframe
driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))

RITM = driver.find_element_by_id('sys_display.sc_task.request_item').get_attribute("value")
print(RITM)
writeToCSV(RITM, 'output.csv')

# Enter comment
inputComment("Test123")

selTicketStatusFromDropdown('Work in Progress')
selTaskTypeFromDropdown('Restock')

submitTicket('false')

# driver.close()
