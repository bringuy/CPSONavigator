
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
cities = ["Kitchener", "Waterloo", "Cambridge"]

def parseAddress(address):
    split_address = address.split()
    city = split_address[0]
    province = split_address[1]
    postal_code = " ".join(split_address[2:])
    return city, province, postal_code    

def trim_list(lst):
    # Find start and end indices
    start_index = lst.index('Location of Practice:') + 1
    end_index = None
    for word in lst:
        if 'This doctor has' in word:
            end_index = lst.index(word)
            break
        elif 'Area(s)' in word:
            end_index = lst.index(word)
            break

    # Get relevant elements
    if end_index != None:
        info_list = lst[start_index:end_index]
        return info_list 
    
def extract_info(lst):

    streetOne = lst[0]
    streetTwo = ''
    inc = 0
    if len(lst) == 5:
        streetTwo = lst[1]
        inc = 1

    try:
        city, province, postal_code = parseAddress(lst[1+inc])
        phone = ""
        fax = ""
        for item in lst[2+inc:]:
            if "Phone:" in item:
                phone = item.replace("Phone: ", "")
            elif "Fax:" in item:
                fax = item.replace("Fax: ", "")

        return ["TBD", streetOne, streetTwo, city, province, postal_code, "Canada", phone, fax]
    except Exception as e:
        print(f"Error occured: {e}")
    else:
        return None    


def main():
    for city in cities:
        # navigate to the form URL
        driver.get("https://doctors.cpso.on.ca/?search=general")

        # Wait for the form to load
        wait = WebDriverWait(driver, 10)
        form = wait.until(EC.presence_of_element_located((By.TAG_NAME, "form")))

        # fill out form
        dropdown = Select(driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl02_CPSO_AllDoctorsSearch_ddCity"))
        dropdown.select_by_visible_text(city)

        is_family_doctor = driver.find_element(By.XPATH,'//label[@for="rdoDocTypeFamly"]').click()

        submit_button = driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl02_CPSO_AllDoctorsSearch_btnSubmit1")
        submit_button.click()

        # Wait for the response page to load
        article_container = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "doctor-search-results")))
        articles = article_container.find_elements(By.TAG_NAME, "article")

        pages = driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl03_CPSO_DoctorSearchResults_lnbLastPage").text

        for page in range(1, int(pages)):
            for row, article in enumerate(articles):
                resp_lines = article.text.splitlines()
                print(resp_lines)

                # get full name
                full_name = resp_lines[0].split("(")[0].strip()
                line_cntr = 0

                cell_vals = trim_list(resp_lines)

                if cell_vals == None:
                    sheet[f"A{(page-1)*10+row+2}"] = "Error when trimming" 
                    sheet[f"B{(page-1)*10+row+2}"] = f"{resp_lines}"
                    continue

                if len(cell_vals) > 5:
                    sheet[f"A{(page-1)*10+row+2}"] = "Too much information can't extract"
                    sheet[f"B{(page-1)*10+row+2}"] = f"{cell_vals}"
                    continue

                cell_vals = extract_info(cell_vals) 

                if cell_vals == None:
                    sheet[f"A{(page-1)*10+row+2}"] = "Error when trimming" 
                    sheet[f"B{(page-1)*10+row+2}"] = f"{resp_lines}"
                    continue

                cell_vals = [full_name] + cell_vals

                for idx, col in enumerate(columns):
                    sheet[f"{col}{(page-1)*10+row+2}"] = cell_vals[idx]
    
                sheet[f"K{(page-1)*10+row+2}"] = f"{cell_vals}"

            if page%5 != 0:
                nextPage = driver.find_element(By.ID, f"p_lt_ctl01_pageplaceholder_p_lt_ctl03_CPSO_DoctorSearchResults_rptPages_ctl0{page%5}_lnbPage")
                nextPage.click()
            else:
                nextGroup = driver.find_element(By.ID, "p_lt_ctl01_pageplaceholder_p_lt_ctl03_CPSO_DoctorSearchResults_lnbNextGroup")
                nextGroup.click() 

            # Wait for the response page to load
            wait = WebDriverWait(driver, 5)
            article_container = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "doctor-search-results")))
            articles = article_container.find_elements(By.TAG_NAME, "article")

    # close the browser
    workbook.save('output.xlsx')
    driver.close()


main()
