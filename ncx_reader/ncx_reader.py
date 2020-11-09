from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import time
import pandas as pd
import os

class ncx_reader:
    def __init__(self,
                 driver= "./chromedriver_85",
                 target_url='https://ncxtools.telkom.co.id/login'):
        self.driver = driver
        self.target_url = target_url
        self.download_folder = "./files/"
        self.sc_number = 0

    def open_target_url(self):
        self.driver = webdriver.Chrome(self.driver)
        if 'browserVersion' in self.driver.capabilities:
            # print("step 1 (load target url).")
            print(f"opening target url: {self.target_url}")
            self.driver.get(self.target_url)
        else:
            print("plese check your chrome version, "
                  "this chrome driver only support version: "+self.driver.capabilities['version'])

    def login(self):
        # print("step 2 (login).")
        keyPressed = input("please login manually first. . . press (y) if done?")
        while keyPressed != "y":
            print("you've presssed {input}")
            keyPressed = input("please login manually first. . . press (y) if done?")
        print("oke, go to the next step")

    def input_scnumber(self, sc_number):
        self.driver.get("https://ncxtools.telkom.co.id/cons/salesOrder")
        self.sc_number = sc_number
        if self.server_error():
            self.input_scnumber(self.sc_number)
        try:
            sleep(1)
            self.driver.find_element_by_xpath('//form[@id="search"]')
            fill_here_input = self.driver.find_element_by_xpath('//input[@name="value"]')
            fill_here_input.send_keys("SC"+str(sc_number))
            fill_here_input.send_keys(u'\ue007')
        except NoSuchElementException:
            pass
            # self.input_scnumber(self.sc_number, max_attempt)
        # cek errors
        if self.not_found_message():
            return 0
        if self.check_error_message():
            self.input_scnumber(sc_number)
        print(f"SC{sc_number}, found!")
        return 1

    def check_error_message(self):
        sleep(1)
        try:
            response = self.driver.find_element_by_xpath('//div[@id="response"]')
            error_message = '''{
        "message": "Server Error"
    }'''
            not_found_message = '''Order Not Found'''
            if error_message in response.text:
                print("print error message found")
                return True
            else:
                return False
        except NoSuchElementException:
            return False

    def not_found_message(self):
        sleep(1)
        try:
            response = self.driver.find_element_by_xpath('//div[@id="response"]')
            not_found_message = '''Order Not Found'''
            if not_found_message in response.text:
                print("External ID not found")
                return True
            else:
                return False
        except NoSuchElementException:
            return False

    def server_error(self):
        sleep(1)
        try:
            response = self.driver.find_element_by_xpath('//div[@class="code"]')
            message = '500'
            if message in response.text:
                print("Server error 500")
                return True
            else:
                return False
        except NoSuchElementException:
            return False

    def access_order_detail(self):
        while self.server_error():
            self.input_scnumber(self.sc_number)
        while self.check_error_message():
            self.input_scnumber(self.sc_number)
        if self.not_found_message():
            print(f"{self.sc_number}: order detail not found")
            return 0
        print(f"SC{self.sc_number}, accessing order detail. ")
        sleep(2)
        try:
            detailButton = self.driver.find_element_by_xpath('//table[@id="dataTable"]/tbody/tr[1]/td[10]')
            if detailButton is not None:
                sleep(2)
                # element = WebDriverWait(self.driver, 20).until(
                #     EC.presence_of_element_located((By.XPATH, '//table[@id="dataTable"]/tbody/tr[1]/td[10]')))
                # self.driver.execute_script("arguments[0].click();", element)
                detailButton.click()
                print("detail button have found, clickable and clicked!")
                sleep(1)
                return 1
        except NoSuchElementException:
            print("detail button NOT found! repeat process.")
            self.input_scnumber(self.sc_number)
            self.access_order_detail()
        return 1

    def get_status(self):
        status = self.driver.find_elements_by_xpath("//div[@id='header']//div[@class='p-4 border-bottom']")
        try:
            li = list(status[0].text.split("\n"))
            return li[li.index("Status") + 2]
        except IndexError:
            return "null"

    def get_submit_flag(self):
        status = self.driver.find_elements_by_xpath("//div[@id='header']//div[@class='p-4 border-bottom']")
        try:
            li = list(status[0].text.split("\n"))
            return li[li.index("Submit flag") + 2]
        except IndexError:
            return "null"

    def get_date_from_str(self, words):
        underscores = 0
        dates = []
        for word in words:
            if word == "_":
                underscores += 1
            if underscores > 3 and len(dates) < 11:
                dates.append(word)
        date_str = ''.join([str(elem) for elem in dates])
        return(date_str[1:])

    def close_browser(self):
        self.driver.close()
        print("browser closed")

if __name__ == '__main__':
    ncx = ncx_reader()
    ncx.open_target_url()
    ncx.login()

    # file to be checked
    pwd = os.getcwd()
    file_name = "report_provi_NCX_FAILED_2020_11_02.xlsx"
    file_date = ncx.get_date_from_str(file_name)
    df = pd.read_excel(pwd+"/ncx_reader/files/"+file_name)

    status_list = []
    submit_flag_list = []
    not_found_list = []

    start_index = 79
    # end_index = 11
    print("there are "+ str(len(df['ORDER_ID'][start_index:]))+ " EXTERNAL ID(s) to be assesed.")
    start = time.time()
    for i, order_id in enumerate(df["ORDER_ID"][start_index:]):
        print(f"Row {start_index+i} : SC{order_id}")
        if ncx.input_scnumber(order_id):
            if ncx.access_order_detail():
                print(f"status      : {ncx.get_status()}")
                print(f"submit flag : {ncx.get_submit_flag()}")
                # status_list.append(ncx.get_status())
                # submit_flag_list.append(ncx.get_submit_flag())
                df.loc[df.index[start_index + i], 'STATUS'] = ncx.get_status()
                df.loc[df.index[start_index + i], 'SUBMIT_FLAG'] = ncx.get_submit_flag()
            else:
                print(f"status      : NOT FOUND")
                print(f"submit flag : NOT FOUND")
                not_found_list.append(order_id)
                df.loc[df.index[start_index + i], 'STATUS'] = 'NOT FOUND'
                df.loc[df.index[start_index + i], 'SUBMIT_FLAG'] = 'NOT FOUND'
        else:
            print(f"status          : NOT FOUND")
            print(f"submit flag     : NOT FOUND")
            not_found_list.append(order_id)
            df.loc[df.index[start_index + i], 'STATUS'] = 'NOT FOUND'
            df.loc[df.index[start_index + i], 'SUBMIT_FLAG'] = 'NOT FOUND'
        print("")
    ncx.close_browser()

    # write excel all
    df.to_excel(pwd+"/ncx_reader/files/"+file_name.split(".")[0]+"_ncx_checked.xlsx")
    end = time.time()
    print(f"done and dusted time lapsed: {(end - start) / 60} minutes")

    # write excel according to where tele group to send
    df_to_kawal_ffdit = df[(df["SUBMIT_FLAG"] == "Y") & (df["STATUS"] != "Cancelled") & (df["STATUS"] != "Completed")]
    df_to_kawal_ncx = df[(df["SUBMIT_FLAG"] == "N") & (df["STATUS"] != "In Progress") & (df["STATUS"] != "Cancelled") & (df["STATUS"] != "Completed")]

    cmds = []
    for index, row in df_to_kawal_ncx.iterrows():
        witel = row['WITEL']
        if (witel == 'KUPANG'):
            witel = 'NTT'
        elif (witel == 'SBY UTARA'):
            witel = 'SURABAYA UTARA'
        else:
            pass
        cmds.append(f"/moban #{witel} #SC{row['ORDER_ID']} #Dorong Push PI")

    for i, command in enumerate(cmds):
        df.loc[df.index[i], 'CMD'] = cmds[i]
        # df_to_kawal_ncx["cmd"] = cmds

    print("writing excel files to be escalated in tele groups")
    with pd.ExcelWriter(pwd+"/ncx_reader/files/"+file_name.split(".")[0]+"_to_ncx_check_kawal_ncx.xlsx") as writer:
        df_to_kawal_ncx.to_excel(writer, sheet_name=file_date)

    with pd.ExcelWriter(pwd+"/ncx_reader/files/"+file_name.split(".")[0]+"_to_ncx_check_ffdit.xlsx") as writer:
        df_to_kawal_ffdit.to_excel(writer, sheet_name=file_date)

    status_list_df = pd.DataFrame(status_list)
    status_list_df.to_excel(pwd+"/ncx_reader/logs/status_list_df_"+file_name)
    submit_flag_list_df = pd.DataFrame(submit_flag_list)
    submit_flag_list_df.to_excel(pwd+"/ncx_reader/logs/submit_flag_df_"+file_name)
    not_found_list_df = pd.DataFrame(not_found_list)
    not_found_list_df.to_excel(pwd+"/ncx_reader/logs/not_found_list_df_"+file_name)