import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup as soup

url = os.environ.get("BANK_URL") # Add your bank's url here

def visit_website(url):
    driver = webdriver.Firefox()
    driver.get(url)
    return driver

if __name__ == "__main__":
    driver = visit_website(url)
