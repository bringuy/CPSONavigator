from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import openpyxl

workbook = openpyxl.Workbook()
sheet = workbook.active
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
columns = [chr(x) for x in range(ord('A'), ord('J')+1)]

with open("data.txt", "r") as f:
    lines = f.readlines()
    full_name, last_name, first_name = None, None, None

    for row, line in enumerate(lines):
        full_name = line.strip()
        if "," in full_name:
            last_name, first_name = full_name.split(",")
            first_name = first_name.strip()
            last_name = last_name.strip()
        else:
            last_name = full_name
            last_name = last_name.strip()
    
        # print(full_name, first_name, last_name)

        # navigate to the form URL
        driver.get("https://doctors.cpso.on.ca/?search=general")

        # Wait for the form to load
        wait = WebDriverWait(driver, 10)
        form = wait.until(EC.presence_of_element_located((By.TAG_NAME, "form")))

        # fill out form
        # is_family_doctor = driver.find_element(By.ID, "rdoDocTypeFamly").click()

        last_name_element = driver.find_element(By.ID, "txtLastName").send_keys(last_name)
        if first_name != None:
            first_name_element = driver.find_element(By.ID, "txtFirstName").send_keys(first_name)

        submit_button = driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl02_CPSO_AllDoctorsSearch_btnSubmit1")
        submit_button.click()

        # Wait for the response page to load
        response = wait.until(EC.presence_of_element_located((By.TAG_NAME, "article")))
        resp_lines = response.text.splitlines()
        print(resp_lines)

        if resp_lines.count("Location of Practice:") > 1:
            sheet[f"K{row+1}"] = "More than one doctor so double check"

        start, end = 0, 0
        for idx, resp_line in enumerate(resp_lines):
            if resp_line == "Location of Practice:":
                start = idx + 1
            elif resp_line == "Area(s) of Specialization:":
                end = idx
        resp_lines = resp_lines[start:end]
        line_cntr = 0

'''
        for idx, col in enumerate(columns):
            cell_val = ""
            if idx == 0:
                cell_val = full_name
            elif idx == 1:
                cell_val = 'TBD'
            elif idx == 2: # street 1 
                cell_val = resp_lines[0] 
            elif idx == 3: # street 2
                if "ON" not in resp_lines[1]:
                    cell_val = resp_lines[1]
            elif idx == 4: # City
                if "ON" not in resp_lines[1]:
                    call_val = resp_lines[2]

            sheet['f{col}{row+1}'] = cell_val 
            '''
    

# close the browser
driver.close()