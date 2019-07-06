from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import csv

# Start global variables
driver = webdriver.Chrome()
saveTicket = False
driver.implicitly_wait(3)


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
    else:
        print("Ticket not saved! Make sure to pass 'true' if you want to actually submit.")


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
            print(row)
            line_count += 1
        print(f'Processed {line_count} lines.')
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
    commentArea = driver.find_element_by_xpath("//textarea[@id='activity-stream-comments-textarea']")
    commentArea.click()
    commentArea.send_keys(commentString)
    print("Typing " + str(commentString))


def search(inputString):
    searchbox = driver.find_element_by_xpath("//*[@id='sysparm_search']")
    searchbox.click()
    searchbox.send_keys(computers[0])
    searchbox.send_keys(Keys.RETURN)

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
# driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Fsc_task.do%3Fsys_id%3D07feba5bdb1677c0d4aa710439961995%26sysparm_view%3Dtext_search%26sysparm_record_target%3Dsc_task%26sysparm_record_row%3D3%26sysparm_record_rows%3D4%26sysparm_record_list%3D123TEXTQUERY321%25253DDT2UA50516QQ")

# Open to service catalog
driver.get("https://ghsprod.service-now.com/nav_to.do?uri=%2Fcatalog_home.do%3Fsysparm_view%3Dcatalog_default")

# Make sure we have the default/top iframe selected
driver.switch_to.default_content()

# Search for the 1st computer
search(computers[0])

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

print("Wait")
time.sleep(5)

# Set ticket state
selTicketStateFromDropdown('Work in Progress')

print("Wait")
time.sleep(5)

# Enter comment
inputComment("Acknowledging device/asset")

print("Wait")
time.sleep(5)

# Save ticket
submitTicket()

print("Wait")
time.sleep(5)

driver.switch_to.default_content()
# Switch to main content iframe
driver.switch_to.frame(driver.find_element_by_class_name("navpage-main-left").find_element_by_xpath(".//iframe"))

# Set ticket state
selTicketStateFromDropdown('Closed Complete')

print("Wait")
time.sleep(5)

# Enter comment
inputComment("Device to be restocked")

print("Wait")
time.sleep(5)

# Set ticket task (Restock, Decommission, or Repair)
selTaskTypeFromDropdown('Restock')

print("Wait")
time.sleep(5)

# Save RITM to CSV
RITM = driver.find_element_by_id('sys_display.sc_task.request_item').get_attribute("value")
print(RITM)
writeToCSV(RITM, 'output.csv')

print("Wait")
time.sleep(5)

# Save ticket
submitTicket()

print("Wait")
time.sleep(5)