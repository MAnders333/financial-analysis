import os
import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup as soup

class FinancialAnalyzer:
    """
    """
    def __init__(self, path=None, sheet_name="Financial Overview"):
        """
        """
        self.sheet_name = sheet_name
        if path != None:
            self.path = str(path)
            try:
                self.book = openpyxl.load_workbook(self.path)
            except FileNotFoundError:
                print("File could not be found. Please provide a valid path.")
            try:
                self.sheet = self.book[sheet_name]
            except:
                print("Worksheet could not be found. Please provide a valid sheet name.")
            self.last_row = self.sheet.max_row
            self.current_month = datetime.date(datetime.now()).month
            if int(self.sheet.cell(self.last_row,2).value) != int(self.current_month):
                self.isUpdated = False
                print("Worksheet needs to be updated.")
            else:
                self.isUpdated = True
                print("Worksheet is up to date.")  
        else:
            self.book = openpyxl.Workbook()
            print("Workbook created.")
            self.sheet = self.book.create_sheet(title=sheet_name, index=0)
            print("Worksheet created.")
            headers = ["year", "month", "cash", "cashflow", "bank", "bank_cashflow", "investments", "investments_total_change", "assets_total", "assets_total_change", "assets_rel_change"]
            for i, header in enumerate(headers):
                self.sheet.cell(1, i+1).value = header
            self.book.save("monthly_budget.xlsx")
            self.isUpdated = False
            print("Worksheet needs to be updated.")

    def scrape_bank_website(self, url, username, password):
        """
        This method needs to be adapted for different bank websites to scrape.
        """
        # Visits the bank website
        self.driver = webdriver.Firefox()
        self.driver.get(url)

        # Waits for a specified element to render to make sure the page is fully loaded
        WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.XPATH, """//*[@id="vrkennungalias"]"""))) # Change the xpath according to the element's xpath of your banks website

        # Writes the username and password into the login form
        element_username = self.driver.find_element_by_xpath("""//*[@id="vrkennungalias"]""")
        element_username.send_keys(username)
        element_password = self.driver.find_element_by_xpath("""//*[@id="pin"]""")
        element_password.send_keys(password)
        
        # Clicks the login button
        login_button = self.driver.find_element_by_xpath("""//*[@id="button_login"]""")
        login_button.click()

        # Waits for a specified element to render to make sure the page is fully loaded
        WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.XPATH, """/html/body/div[2]/div[3]/div[2]/div[1]/div/div/div/div/div/div/h1""")))

        # Reads out balances in the bank accounts
        self.balance1 = float(self.driver.find_element_by_xpath("""/html/body/div[2]/div[3]/div[2]/div[1]/div/form/div[2]/table/tbody[1]/tr[2]/td[4]""").text.strip().split(" ")[0].replace(".", "").replace(",", "."))
        self.balance2 = float(self.driver.find_element_by_xpath("""/html/body/div[2]/div[3]/div[2]/div[1]/div/form/div[2]/table/tbody[2]/tr[2]/td[4]""").text.strip().split(" ")[0].replace(".", "").replace(",", "."))

        # Quits browser
        self.driver.quit()
        
    def update_workbook(self):
        pass




if __name__ == "__main__":
    path = "/Users/marcanders/Desktop/monthly_budget.xlsx" # Add the path to your Excel file here
    url = os.environ.get("BANK_URL") # Add your bank's url here
    username = os.environ.get("BK_UN") # Add your bank's username here
    password = os.environ.get("BK_PW") # Add your bank's password here
    
    analyzer = FinancialAnalyzer(path=path)
    if analyzer.isUpdated == False:
        analyzer.scrape_bank_website(url, username, password) 
        analyzer.update_workbook()

    







