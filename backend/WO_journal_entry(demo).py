#!/bin/env python3.7
__author__ = "Tim Zong"

from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
import getpass
from webdriver_manager.chrome import ChromeDriverManager

class WOJournalEntry():
    def setup_method(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        # options.add_argument("--no-sandbox")
        options.add_argument("--remote-debugging-port=9222")
        # options.add_argument("--headless")

        self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        self.driver.maximize_window()

        self.vars = {}

    def teardown_method(self):
        self.driver.quit()

    def run(self):

        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name("RPA 'start WO Journal Entry'-26c277e69666.json", scope)
        client = gspread.authorize(creds)
        google_sheet = client.open('Workorder Journals (Responses)')
        approved_sheet = google_sheet.worksheet("Approved")
        records = approved_sheet.get_all_records()
        count =0

        self.login()
        for i in range(len(approved_sheet.col_values(16)),len(records)+1): # check from column P
            cell_list = approved_sheet.range("B{}:J{}".format(i + 1, i + 1))
            description = cell_list[0].value
            subledger = cell_list[1].value
            totalcost = cell_list[2].value
            WO = cell_list[3].value
            phase = cell_list[4].value
            offset_SC = cell_list[5].value
            offset_acct = cell_list[6].value
            charge_SC = cell_list[7].value
            charge_acct = cell_list[8].value
            # description = approved_sheet.cell(i+1,2).value # column B
            # subledger = approved_sheet.cell(i+1,3).value # column C
            if subledger=="Overhead":
                subledger = "Equipment"
            # totalcost = approved_sheet.cell(i+1,4).value # column D
            # WO = approved_sheet.cell(i+1,5).value # column E
            # phase = approved_sheet.cell(i+1,6).value # column F
            # offset_SC = approved_sheet.cell(i+1,7).value # column G
            # offset_acct = approved_sheet.cell(i+1,8).value # column H
            # charge_SC = approved_sheet.cell(i+1,9).value # column I
            # charge_acct = approved_sheet.cell(i+1,10).value # column J

            tranx_no,error_text = self.writing_data_to_aim(description,subledger,totalcost,WO,phase,
                                                offset_SC,offset_acct,charge_SC,charge_acct)
            if error_text=="": # No error, sucessfully saved in AiM
                approved_sheet.update_cell(i+1,16,tranx_no) # column P
                current_time = str(datetime.now())
                approved_sheet.update_cell(i+1,17,current_time)# column Q
                count += 1
            else:
                approved_sheet.update_cell(i+1,18,error_text)# column R

        print ("{0} of rows have been logged into AiM!".format(count))
        input("Press Enter to continue...")

    def login(self):
        self.driver.get("https://www.aimdemo.ualberta.ca/fmax/login")
        # self.driver.set_window_size(1900, 1020)
        username = input('Enter your username: ')
        password = getpass.getpass('Enter your password : ')
        self.driver.find_element(By.ID, "username").send_keys(username)
        self.driver.find_element(By.ID, "password").send_keys(password)
        self.driver.find_element(By.ID, "login").click()
        self.driver.find_element(By.ID, "mainForm:menuListMain:FINANCE").click()
        self.driver.find_element(By.ID, "mainForm:menuListMain:search_WO_JOURNAL_VIEW").click()
        self.driver.find_element(By.ID, "mainForm:buttonPanel:executeSearch").click()

    def writing_data_to_aim(self,desc,sub,cost,wo,ph,off_sc,off_acct,charge_sc,charge_acct):
        tranx_no = ""
        error_message = ""

        self.driver.find_element(By.ID, "mainForm:buttonPanel:new").click()
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:ae_s_wo_journal_e_description").send_keys(desc)
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:sublegerValue").click()
        dropdown = self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:sublegerValue")
        dropdown.find_element(By.XPATH, "//option[. = '{0}']".format(sub)).click()
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:adjAmountValue").send_keys(cost)
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:workOrderPhaseZoom:level0").send_keys(wo)
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:workOrderPhaseZoom:level1").send_keys(ph)
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:accountCodeZoom:acctZoom0").send_keys(off_sc)
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:accountCodeZoom:acctZoom1").send_keys(off_acct)

        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:detailAccountList:addChargeAccountsButton").click()
        self.driver.find_element(By.ID, "mainForm:buttonPanel:zoomNext").click()

        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_ACCOUNT_EDIT_content:accountCodeZoom:level1").send_keys(charge_sc)
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_ACCOUNT_EDIT_content:accountCodeZoom:level2").send_keys(charge_acct)
        self.driver.find_element(By.ID, "mainForm:buttonPanel:done").click()
        try:
            self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:postFlagValue").click()
        except:
            error_message = self.driver.find_element(By.ID,"mainForm:WO_JOURNAL_ACCOUNT_EDIT_content:messages").text
            self.driver.find_element(By.ID,"mainForm:buttonPanel:cancel").click()
            self.driver.find_element(By.ID, "mainForm:buttonPanel:cancel").click()
            return (tranx_no,error_message)
        dropdown = self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:postFlagValue")
        dropdown.find_element(By.XPATH, "//option[. = 'Yes']").click()
        self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:postFlagValue").click()

        tranx_no = self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:ae_s_wo_journal_e_transaction_num").text
        self.driver.find_element(By.ID, "mainForm:buttonPanel:save").click()
        try:
            new_button = self.driver.find_element(By.ID, "mainForm:buttonPanel:new")
        except:
            error_message = self.driver.find_element(By.ID, "mainForm:WO_JOURNAL_EDIT_content:messages").text
            self.driver.find_element(By.ID, "mainForm:buttonPanel:cancel").click()
        return (tranx_no,error_message)



if __name__ == '__main__':
    entry = WOJournalEntry()
    entry.setup_method()
    entry.run()
    entry.teardown_method()