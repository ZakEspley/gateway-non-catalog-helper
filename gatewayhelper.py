import re
import os
import sys

#### from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QFileDialog, QApplication
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QThread
from ui_gatewayhelper import Ui_MainWindow

import pandas as pd
from selenium import webdriver as wd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import ui
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchWindowException
import platform
import json


class GatewayHelperApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(GatewayHelperApp, self).__init__()

        if ("windows" in platform.system().lower()):
            self.driver_exec = "chromedriver.exe"
            self.name = "GatewayHelper.exe"
        else:
            self.driver_exec = "chromedriver"
            self.name = "GatewayHelper.app"

        if getattr(sys, 'frozen', False):
            self.tempPath = sys._MEIPASS
            self.application_path = os.path.dirname(sys.executable)
            self.resourcePath = os.path.join(self.application_path, "..", "Resources")
            self.driverPath = os.path.join(self.application_path, self.driver_exec)
        elif __file__:
            self.application_path = os.path.dirname(__file__)
            self.resourcePath = self.application_path
            self.tempPath = self.application_path
            self.driverPath = os.path.join(os.getcwd(), self.driver_exec)

        with open(os.path.join(self.resourcePath, "settings.json")) as file:
            self.settings = json.load(file)

        if len(self.settings["lastPath"]) == 0:
            self.lastPath = os.path.expanduser("~/Documents")
        else:
            self.lastPath = self.settings["lastPath"]

        self.browser = self.settings["browser"]
        self.icon = QIcon( os.path.join(self.resourcePath, "icon.icns") )
        self.cwd = os.getcwd()
        self.filePath = None
        self.vendor = None
        self.setWindowIcon(self.icon)
        self.gatewayLink = "https://sso.ucsb.edu:8443/cas/login?service=https%3a%2f%2fgateway.procurement.ucsb.edu%2fr%2f"
        self.setupUi(self)
        self.browserThread = QThread()
    
    def run(self):
        self.browserThread = driverThread(self, self.driverPath, self.filePath, self.vendor)
        self.browserThread.start()

    def openFile(self):
        self.filePath, fileTypes = QFileDialog.getOpenFileName(self, "Select Cart Excel File", self.lastPath, filter="Excel Files (*.xlsx *xls)")
        self.cartLineEdit.setText(self.filePath)
        if (len(self.filePath) > 0):
            self.lastPath = os.path.dirname(self.filePath)
    
    def updateVendor(self):
        self.vendor = self.vendorLineEdit.text()

    def showAbout(self):
        QMessageBox.about(self, "About", '<qt><p>Created by Zak Espley.</p>  <p>zespley@physics.ucsb.edu</p>  <div>Icons made by <a href="https://www.flaticon.com/authors/twitter" title="Twitter">Twitter</a> from <a href="https://www.flaticon.com/" title="Flaticon">www.flaticon.com</a> is licensed by <a href="http://creativecommons.org/licenses/by/3.0/" title="Creative Commons BY 3.0" target="_blank">CC 3.0 BY</a></div></qt>')

    def cancel(self):
        if (self.browserThread.isRunning):
            self.browserThread.quit()
        try:
            self.browserThread.driver.quit()
        except:
            pass
        self.__updateSettings(self.browser, self.lastPath)
        self.close()

    def __updateSettings(self, browser, lastpath):
        self.settings["browser"] = browser
        self.settings["lastPath"] = lastpath
        with open(os.path.join(self.resourcePath, "settings.json"), "w") as file:
            json.dump(self.settings, file)
        
    
    def closeEvent(self, event):
        self.__updateSettings(self.browser, self.lastPath)
        if (self.browserThread.isRunning):
            self.browserThread.quit()
        try:
            self.browserThread.driver.quit()
        except:
            pass
        event.accept()
    
class driverThread(QThread):

    def __init__(self, parent, driverPath, filePath, vendor):
        super(driverThread, self).__init__()
        self.parent = parent
        self.driverPath = driverPath
        self.filePath = filePath
        self.vendor = vendor
    
    def run(self):

        if len(self.vendor) <= 0:
            QMessageBox.warning(self.parent, 'Vendor Warning', 'You need to enter a vendor!')
            return None

        try:
            self.df = pd.read_excel(self.filePath)
        except (FileNotFoundError, ValueError):
            QMessageBox.warning(self.parent,'Cart Warning', 'You need to enter a proper file!')
            return None
        regex = re.compile('[$]')

        try:
            
            self.driver = wd.Chrome(executable_path=self.driverPath)

            self.driver.get(self.parent.gatewayLink)

            wait = ui.WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, "FormSkuSearchLink")))
            link = self.driver.find_element_by_id("FormSkuSearchLink")
            link.click()
            ui.WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "ModalPopupIframe")))
            self.driver.switch_to.frame(self.driver.find_element_by_id("ModalPopupIframe"))
            ui.WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "newSupplierSearchId")))

            supplier_input = self.driver.find_element_by_id("newSupplierSearchId")
            supplier_input.clear()
            supplier_input.send_keys(self.vendor)
            ui.WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//li[@class='yui-ac-highlight']")))
            supplier_input.send_keys(Keys.RETURN)

            fill_in_data = {
                "Description":(By.ID, "textAreaID_11"),
                "Non-Catalog #": (By.NAME, "NonCatCatalogNumber"),
                "Qty":(By.NAME, "NonCatQuantity"),
                "Price (USD)":(By.NAME, "NonCatUnitPrice")
            }
            total_rows = self.df["Non-Catalog #"].count()
            for index, row in self.df.iterrows():

                for key, value in fill_in_data.items():
                    elem = self.driver.find_element(*value)
                    elem.send_keys(regex.sub('',str(row[key])))

                if index == total_rows-1:
                    button = self.driver.find_element_by_xpath("//input[@type='button' and @value='Save and Close']")
                    button.click()
                else:
                    button = self.driver.find_element_by_xpath("//input[@type='button' and @value='Save and Add Another']")
                    button.click()

            self.quit()

        except TimeoutException as e:
            QMessageBox.critical(self.parent, 'Network Issues', 'You are either having network problems or took to long to do something. \r\n\r\n Exiting.')
            self.driver.close()
            self.quit()
        
        except NoSuchWindowException:
            pass

if __name__=="__main__":
    app = QApplication(sys.argv)
    window = GatewayHelperApp()
    window.setWindowTitle("Gateway Helper")
    window.show()
    sys.exit(app.exec_())