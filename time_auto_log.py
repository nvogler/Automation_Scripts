import sys, os, time, datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# Temp add to PATH
os.environ['PATH'] += ';H:\\Scripts\\TimeScript'

# Identify charge codes
cost_center = ""
charge_code = ""

# Open FireFox
driver = webdriver.Firefox()

# Load and Wait
driver.get("TIME WEB APPLICATION LOCATION")

print("Waiting...")
time.sleep(3)
print("Done.")

# Navigate to Time Keeping Tab and Wait
mainpage = driver.find_element_by_id("tabIcon2")
mainpage.click()
print("Waiting...")
time.sleep(3)
print("Done.")

# Diving into the IFrames
javaTableOuter = driver.find_element_by_id("contentAreaFrame")
driver.switch_to_frame(javaTableOuter)

# Further into the IFrames
javaTableInner = driver.find_element_by_id("ivuFrm_page0ivu0")
driver.switch_to_frame(javaTableInner)

# Enter Network
## Determine next free row
rowNum = 0
newNetworkCell = ""
for i in range(2,99):
	id_net = "aaaa.RecordTimeView.NetworkWLInputField." + str(i)
	id_wcr = "aaaa.RecordTimeView.WorkcenterWLInputField." + str(i)
	if driver.find_element_by_name(id_net).get_attribute("value") == "":
		## Input Data
		driver.find_element_by_name(id_net).send_keys(charge_code)
		driver.find_element_by_name(id_wcr).send_keys(cost_center)
		# Stash
		rowNum = i
		newNetworkCell = id_net
		break
		
# Enter Time
## Determine date from system time
date = (str(datetime.date.today())[5:]).replace('-', '/')

# Find Column
col = driver.find_element_by_xpath("//span[text()='" + date + "']")

# Pull ID
day = str(2)
newTimeCell = col.get_attribute("id").split('_')[0] + "_editor." + day

# Input Data
driver.find_element_by_id(newTimeCell).clear()
driver.find_element_by_id(newTimeCell).send_keys("9")

# Goto Review
driver.find_element_by_id("aaaa.RecordTimeView.ReviewButton").click()

print("Waiting...")
time.sleep(2)
print("Done.")
# Confirm Changes
driver.find_element_by_id("aaaa.ReviewNConfirmTimecardView.SaveButton").click()

# Close Browser
print("Waiting...")
time.sleep(2)
driver.close()