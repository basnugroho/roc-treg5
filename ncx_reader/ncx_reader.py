from selenium import webdriver
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

    def open_target_url(self):
        self.driver = webdriver.Chrome(self.driver)
        if 'browserVersion' in self.driver.capabilities:
            print("step 1 (load target url).")
            print(f"opening target url: {self.target_url}")
            self.driver.get(self.target_url)
        else:
            print("plese check your chrome version, "
                  "this chrome driver only support version: "+self.driver.capabilities['version'])

    def login(self):
        print("step 2 (login).")
        keyPressed = input("please login manually first. . . press (y) if done?")
        while keyPressed != "y":
            print("you've presssed {input}")
            keyPressed = input("please login manually first. . . press (y) if done?")
        print("oke, go to the next step")
        self.driver.get("https://ncxtools.telkom.co.id/cons/salesOrder")

    def input_scnumber(self, sc_number):
        print("step 3 (input SC).")
        self.driver.find_element_by_xpath('//form[@id="search"]')
        fill_here_input = self.driver.find_element_by_xpath('//input[@name="value"]')
        fill_here_input.send_keys("SC"+str(sc_number))
        fill_here_input.send_keys(u'\ue007')
        self

    def access_order_detail(self):
        sleep(1)
        print("step 4 (access order detail).")
        self.driver.find_element_by_xpath('//table[@id="dataTable"]/tbody/tr[1]/td[10]').click()
        sleep(1)

    def get_status(self):
        status = self.driver.find_elements_by_xpath("//div[@id='header']//div[@class='p-4 border-bottom']")
        li = list(status[0].text.split("\n"))
        return li[li.index("Status") + 2]

    def get_submit_flag(self):
        status = self.driver.find_elements_by_xpath("//div[@id='header']//div[@class='p-4 border-bottom']")
        li = list(status[0].text.split("\n"))
        return li[li.index("Submit flag") + 2]

    def close_browser(self):
        self.driver.close()
        print("browser closed")

if __name__ == '__main__':
    ncx = ncx_reader()
    ncx.open_target_url()
    ncx.login()
    ncx.input_scnumber(503564335)
    ncx.access_order_detail()
    print(f"status: {ncx.get_status()}")
    print(f"submit flag: {ncx.get_submit_flag()}")
    ncx.close_browser()

