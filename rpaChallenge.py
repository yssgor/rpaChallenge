from selenium import webdriver
from selenium.webdriver.common.by import By
import time, sys
import pandas as pd

#setting webdriver
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
"download.default_directory": sys.path[0],
"download.prompt_for_download": False,
"download.directory_upgrade": True,
"safebrowsing.enabled": True
})
driver = webdriver.Chrome(options=options)

def download_file():
    link_download = driver.find_element(By.PARTIAL_LINK_TEXT, "DOWNLOAD EXCEL")
    link_download.click()
    time.sleep(2)
    inject_form()

def inject_form():
    try:
        file = pd.read_excel("challenge.xlsx", index_col=0)
    except FileNotFoundError:
        print("File not found.")

    for i, row in file.iterrows():
        first_name = row.name
        last_name = row.iloc[0]
        company_name = row.iloc[1]
        role_in_company = row.iloc[2]
        address = row.iloc[3]
        email = row.iloc[4]
        phone_number = row.iloc[5]

        input = driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelFirstName']")
        input.send_keys(first_name)

        input = driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelLastName']")
        input.send_keys(last_name)

        input = driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelCompanyName']")
        input.send_keys(company_name)

        input = driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelRole']")
        input.send_keys(role_in_company)

        input = driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelAddress']")
        input.send_keys(address)

        input = driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelEmail']")
        input.send_keys(email)

        input = driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelPhone']")
        input.send_keys(phone_number)

        time.sleep(1)

        btn_submit = driver.find_element(By.XPATH, "//input[@value='Submit']")
        btn_submit.click()

def start_challenge():
    driver.get('http://www.rpachallenge.com/')
    btn_start = driver.find_element(By.CLASS_NAME, 'waves-effect.col.s12.m12.l12.btn-large.uiColorButton')
    btn_start.click()
    download_file()

if __name__ == '__main__':
    start_challenge()
    time.sleep(2)