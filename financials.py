import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup as soup
import openpyxl

url = os.environ.get("BANK_URL") # Add your bank's url here
username = os.environ.get("BK_UN") # Add your bank's username here
pw = os.environ.get("BK_PW") # Add your bank's password here

def visit_website(url):
    driver = webdriver.Firefox()
    driver.get(url)
    return driver

def wait_until_element_visible(driver, xpath):
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, xpath)))

def return_element(driver, xpath):
    element = driver.find_element_by_xpath(xpath)
    return element

def return_text_of_element(driver, xpath):
    element = driver.find_element_by_xpath(xpath)
    text = element.text
    return text

def quit_browser(driver):
    driver.quit()

if __name__ == "__main__":
    driver = visit_website(url)
    wait_until_element_visible(driver, """//*[@id="vrkennungalias"]""") # Add the xpath of your bank's website for the user input here
    element_user = return_element(driver, """//*[@id="vrkennungalias"]""") # Add the xpath of your bank's website for the user input here
    element_user.send_keys(username)
    element_password = return_element(driver, """//*[@id="pin"]""") # Add the xpath of your bank's website for the pw input here
    element_password.send_keys(pw)
    login_button = return_element(driver, """//*[@id="button_login"]""")
    login_button.click()
    wait_until_element_visible(driver, """/html/body/div[2]/div[3]/div[2]/div[1]/div/div/div/div/div/div/h1""")
    balance1 = return_text_of_element(driver, """/html/body/div[2]/div[3]/div[2]/div[1]/div/form/div[2]/table/tbody[1]/tr[2]/td[4]""").strip()
    balance2 = return_text_of_element(driver, """/html/body/div[2]/div[3]/div[2]/div[1]/div/form/div[2]/table/tbody[2]/tr[2]/td[4]""").strip()
    quit_browser(driver)







