from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import openpyxl

from helper import parseAddress, trim_list, extract_info

# globals
workbook = openpyxl.Workbook()
sheet = workbook.active
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
columns = [chr(x) for x in range(ord('A'), ord('J')+1)]
cities = ["Kitchener", "Waterloo", "Cambridge"]
cpso_url = "https://doctors.cpso.on.ca/?search=general"

def main():
    pageOffset = 0

    for city in cities:
        # navigate to the form URL
        driver.get(cpso_url)

        # Wait for the form to load
        wait = WebDriverWait(driver, 10)
        form = wait.until(EC.presence_of_element_located((By.TAG_NAME, "form")))

        # fill out form
        dropdown = Select(driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl02_CPSO_AllDoctorsSearch_ddCity"))
        dropdown.select_by_visible_text(city)
        is_family_doctor = driver.find_element(By.XPATH,'//label[@for="rdoDocTypeFamly"]').click()

        # submit form
        submit_button = driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl02_CPSO_AllDoctorsSearch_btnSubmit1")
        submit_button.click()

        # Wait for the response page to load
        article_container = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "doctor-search-results")))
        articles = article_container.find_elements(By.TAG_NAME, "article")

        pages = driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl03_CPSO_DoctorSearchResults_lnbLastPage").text

        for page in range(1, int(pages)):
            for row, article in enumerate(articles):
                resp_lines = article.text.splitlines()

                # get full name + main contact
                full_name = resp_lines[0].split("(")[0].strip()
                parts = full_name.split()
                first_name, last_name = parts[1][0], parts[0]
                main_contact = "Dr. " + last_name + ", " + first_name + "."

                # remove unnecessary info
                cell_vals = trim_list(resp_lines)

                # error checking
                if cell_vals == None:
                    sheet[f"A{pageOffset+(page-1)*10+row+3}"] = "Error when trimming" 
                    sheet[f"B{pageOffset+(page-1)*10+row+3}"] = f"{resp_lines}"
                    continue
                elif len(cell_vals) > 5:
                    sheet[f"A{pageOffset+(page-1)*10+row+3}"] = "Too much information can't extract"
                    sheet[f"B{pageOffset+(page-1)*10+row+3}"] = f"{cell_vals}"
                    continue

                # get information from lines
                cell_vals = extract_info(cell_vals) 

                # error checking
                if cell_vals == None:
                    sheet[f"A{pageOffset+(page-1)*10+row+3}"] = "Error when trimming" 
                    sheet[f"B{pageOffset+(page-1)*10+row+3}"] = f"{resp_lines}"
                    continue

                cell_vals = [full_name, main_contact] + cell_vals

                # update excel sheet row
                for idx, col in enumerate(columns):
                    sheet[f"{col}{pageOffset+(page-1)*10+row+3}"] = cell_vals[idx]
    
                sheet[f"K{pageOffset + (page-1)*10+row+3}"] = f"{cell_vals}"

            # go to next group of pages every five
            if page%5 != 0:
                nextPage = driver.find_element(By.ID, f"p_lt_ctl01_pageplaceholder_p_lt_ctl03_CPSO_DoctorSearchResults_rptPages_ctl0{page%5}_lnbPage")
                nextPage.click()
            else:
                nextGroup = driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl03_CPSO_DoctorSearchResults_lnbNextGroup")
                nextGroup.click() 

            # Wait for the new response page to load
            wait = WebDriverWait(driver, 5)
            article_container = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "doctor-search-results")))
            articles = article_container.find_elements(By.TAG_NAME, "article")

        pageOffset += (int(pages) * 10) # avoid overlap
    # close the browser
    workbook.save('GP_Demographics.xlsx')
    driver.close()


main()
