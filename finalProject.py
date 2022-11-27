from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time

# Vehicle_Num = 'AY65 WRW'
Vehicle_Num = input("Enter the Vehicle Number: ")

options = Options()
options.headless = True
driver = webdriver.Chrome(options=options)

print("Working...")


    # GET URL
driver.get("https://vehicleenquiry.service.gov.uk/")
time.sleep(3)

try: 

    # Entering the Vehicle Number
    driver.find_element(By.XPATH, './/*[@id="wizard_vehicle_enquiry_capture_vrn_vrn"]').send_keys(Vehicle_Num)
    # driver.find_element(By.ID, "lname").send_keys(Vehicle_Num)
    driver.find_element(By.XPATH, """.//*[@id="submit_vrn_button"]""").click()
    time.sleep(7)

    # YES Button
    driver.find_element(By.XPATH, """.//*[@id="yes-vehicle-confirm"]""").click()
    time.sleep(7)
    # Clickking continue
    driver.find_element(By.XPATH, """.//*[@id="capture_confirm_button"]""").click()
    time.sleep(7)

    # GETTING Data
    datas = driver.find_elements(By.ID, "main-content")

    vehNum = ''
    currTaxStat = ''
    currMOTStat = ''
    taxDue = ''
    tdate = ''
    expire = ''
    eDate = ''

    print("Printing Data...")
    time.sleep(3)

    for data in datas:
        vNum = data.find_element(By.CLASS_NAME, "reg-mark").text
        vehNum = vNum

        # print("###########################################")
        headStat = driver.find_elements(By.TAG_NAME, 'h2')
        currTaxStat = headStat[1].text
        currMOTStat = headStat[2].text

        # print("###########################################")
        minStat = data.find_elements(By.TAG_NAME, "strong")
        taxDue = minStat[0].text
        tdate = minStat[1].text
        expire = minStat[2].text
        edate = minStat[3].text


        print(vehNum)
        print(currTaxStat)
        print(taxDue)
        print(tdate)
        print(currMOTStat)
        print(expire)
        print(edate)

    book = load_workbook('data.xlsx')

    sheet = book.active

    rows = [(vehNum, currTaxStat, tdate, currMOTStat, edate)]



    for row in rows:
        sheet.append(row)


    book.save("data.xlsx")

    print("Succesfully Scraped and Appended")

except Exception:
    print("Server Error or Invalid Number")



