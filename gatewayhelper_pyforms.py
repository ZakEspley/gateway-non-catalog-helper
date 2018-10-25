import re
import os
import sys
# import pyforms
import pyforms.settings
from pyforms.basewidget import BaseWidget
from pyforms.controls import ControlFile, ControlText, ControlButton
import pandas as pd
from selenium import webdriver as wd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import ui
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

class  App(BaseWidget):
    def __init__(self, *args, **kwargs):
        super().__init__("Gateway Helper")

        self.set_margin(20)

        if getattr(sys, 'frozen', False):
            self.application_path = os.path.dirname(sys.executable)
        elif __file__:
            self.application_path = os.path.dirname(__file__)

        # cwd = os.getcwd()
        self.cwd = self.application_path
        # print()
        # print(cwd)
        # print()
        # self.cartsDir = os.path.join(self.cwd, "Carts")

        os.environ["PATH"] += os.pathsep + self.cwd


        self._vendor = ControlText('Vendor (Use exactly what Gateway does)')
        self._cart = ControlFile('Cart Excel File')
        self._runbutton = ControlButton('Run')
        self._cancelbutton = ControlButton('Cancel')
        
        self._vendor.changed_event = self.__vendorSelectEvent
        self._cart.changed_event = self.__cartSelectEvent

        self._runbutton.value = self.__runEvent
        self._cancelbutton.value = self.__cancelEvent

        self._formset = [
            ('_vendor'),
            '_cart',
            ("_runbutton", '_cancelbutton')
        ]

        self.fileOkay = False
        self.vendorOkay = False
        self.df = pd.DataFrame()

    def __runEvent(self):

        if len(self._vendor.value) <= 0:
            self.warning("You need to enter a vendor!", "Vendor Warning")
            return None

        try:
            self.df = pd.read_excel(self._cart.value)
        except FileNotFoundError:
            self.alert("FILE NOT FOUND")
            self.warning("You need to enter a proper file!", "Cart Warning")
            return None
        regex = re.compile('[$]')


        

        try:
            
            driver = wd.Chrome()

            driver.get("https://sso.ucsb.edu:8443/cas/login?service=https%3a%2f%2fgateway.procurement.ucsb.edu%2fr%2f")

            wait = ui.WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "FormSkuSearchLink")))
            link = driver.find_element_by_id("FormSkuSearchLink")
            link.click()
            ui.WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ModalPopupIframe")))
            driver.switch_to.frame(driver.find_element_by_id("ModalPopupIframe"))
            ui.WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "newSupplierSearchId")))

            supplier_input = driver.find_element_by_id("newSupplierSearchId")
            supplier_input.clear()
            supplier_input.send_keys(self._vendor.value)
            ui.WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//li[@class='yui-ac-highlight']")))
            supplier_input.send_keys(Keys.RETURN)

            fill_in_data = {
                "Description":(By.ID, "textAreaID_11"),
                "Non-Catalog #": (By.NAME, "NonCatCatalogNumber"),
                "Qty":(By.NAME, "NonCatQuantity"),
                "Price (USD)":(By.NAME, "NonCatUnitPrice")
            }
            total_rows = self.df["Non-Catalog #"].count()
            # print()
            # print("Total Rows:", total_rows)
            # print()
            for index, row in self.df.iterrows():

                # print(index)
                for key, value in fill_in_data.items():
                    elem = driver.find_element(*value)
                    elem.send_keys(regex.sub('',str(row[key])))

                if index == total_rows-1:
                    button = driver.find_element_by_xpath("//input[@type='button' and @value='Save and Close']")
                    button.click()
                else:
                    button = driver.find_element_by_xpath("//input[@type='button' and @value='Save and Add Another']")
                    button.click()

        except TimeoutException as e:
            self.alert("Network Issue", "You are either having network problems, or took to long to do something.")
            self.close()
            sys.exit()

    
    def __cancelEvent(self):
        self.close()
        sys.exit()
    
    def __cartSelectEvent(self):
        pass

    
    def __vendorSelectEvent(self):
        pass



if __name__=="__main__":
    from pyforms import start_app
    start_app(App, geometry=(200,200,500,100))