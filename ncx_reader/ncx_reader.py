from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from time import sleep
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
        print("step 3 (input SC).")
        try:
            self.driver.find_element_by_xpath('//form[@id="search"]')
            fill_here_input = self.driver.find_element_by_xpath('//input[@name="value"]')
            fill_here_input.send_keys("SC"+str(sc_number))
            fill_here_input.send_keys(u'\ue007')
        except NoSuchElementException:
            self.input_scnumber(self.sc_number)

        # cek error { "message": "Server Error" }
        sleep(1)
        server_error = self.check_error_message()
        while server_error:
            self.input_scnumber(sc_number)

    def check_error_message(self):
        response = self.driver.find_element_by_xpath('//div[@id="response"]')
        error_message = '''{
    "message": "Server Error"
}'''
        if error_message in response.text:
            print("print error message found")
            return True
        return False

    def access_order_detail(self):
        sleep(1)
        print("step 4 (access order detail).")
        server_error = self.check_error_message()
        while server_error:
            self.input_scnumber(self.sc_number)
        #why sometime error ?
        try:
            self.driver.find_element_by_xpath('//table[@id="dataTable"]/tbody/tr[1]/td[10]').click()
        except NoSuchElementException:
            self.input_scnumber(self.sc_number)
            self.access_order_detail()
        sleep(1)

    def get_status(self):
        status = self.driver.find_elements_by_xpath("//div[@id='header']//div[@class='p-4 border-bottom']")
        try:
            li = list(status[0].text.split("\n"))
            return li[li.index("Status") + 2]
        except IndexError:
            return ""

    def get_submit_flag(self):
        status = self.driver.find_elements_by_xpath("//div[@id='header']//div[@class='p-4 border-bottom']")
        try:
            li = list(status[0].text.split("\n"))
            return li[li.index("Submit flag") + 2]
        except IndexError:
            return ""

    def close_browser(self):
        self.driver.close()
        print("browser closed")

if __name__ == '__main__':
    ncx = ncx_reader()
    ncx.open_target_url()
    ncx.login()

    pwd = os.getcwd()
    ncx_create_file_name = "report_provi_NCX_CREATE_2020_10_27.xlsx"
    ncx_failed_file_name = "report_provi_NCX_FAILED_2020_10_27.xlsx"
    df = pd.read_excel(pwd+"/ncx_reader/files/"+ncx_failed_file_name)
    status_list = []
    submit_flag_list = []
    print(df["ORDER_ID"])

    for i, order_id in enumerate(df["ORDER_ID"]):
        print(f"order id no. {i+1}: {order_id}")
        ncx.input_scnumber(order_id)
        ncx.access_order_detail()
        print(f"status: {ncx.get_status()}")
        status_list.append(ncx.get_status())
        print(f"submit flag: {ncx.get_submit_flag()}")
        submit_flag_list.append(ncx.get_submit_flag())
    ncx.close_browser()

    # add new column and data
    df["STATUS"] = status_list
    df["SUBMIT_FLAG"] = submit_flag_list

    # write excel according to where tele group to send
    yesterday_for_sheet_name = "2020_10_27"
    date_for_sheet_name = "2020_10_27"
    df_to_kawal_ncx = df[(df["SUBMIT_FLAG"] == "N") & (df["STATUS"] != "Submitted")]
    df_to_kawal_ffdit = df[(df["SUBMIT_FLAG"] == "Y") & (df["STATUS"] != "Submitted")]

    print("writing excel files")
    with pd.ExcelWriter(pwd+"/ncx_reader/files/"+ncx_failed_file_name+"_ncx_check_kawal_ncx.xlsx") as writer:
        df_to_kawal_ncx.to_excel(writer, sheet_name=date_for_sheet_name)

    with pd.ExcelWriter(pwd+"/ncx_reader/files/"+ncx_failed_file_name+"_ncx_check_ffdit.xlsx") as writer:
        df_to_kawal_ffdit.to_excel(writer, sheet_name=date_for_sheet_name)
    print("done")

