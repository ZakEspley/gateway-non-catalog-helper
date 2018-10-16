from tkinter import filedialog, simpledialog, messagebox
from tkinter import Tk
import pandas as pd
from selenium import webdriver as wd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import ui
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import time
import re
import os
import sys

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

# cwd = os.getcwd()
cwd = application_path
# print()
# print(cwd)
# print()
cartsDir = os.path.join(cwd, "Carts")

os.environ["PATH"] += os.pathsep + cwd
tk = Tk()
tk.withdraw()
tk.update()
while True:
    vendor = simpledialog.askstring("Vendor Name", "Which Vendor are you looking to upload to. Use exactly waht Gateway uses if you know.")
    if vendor is None:
        sys.exit()
    elif vendor != '':
        break
    else:
        messagebox.showwarning("No Vendor", "You need to enter in a vendor")

while True:
    file_name = filedialog.askopenfilename(title="Select a BOM Excel to upload", initialdir=cartsDir ,filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_name is None:
        sys.exit()
    elif file_name != '':
        break
    else:
        messagebox.showwarning("No File", "Sorry I didn't see a file.")
tk.deiconify()

regex = re.compile('[$]')


try:
    df = pd.read_excel(file_name)
except FileNotFoundError:
    messagebox.showerror("File Not Found", "Can't file that file. Exiting.")
    sys.exit()

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
    supplier_input.send_keys(vendor)
    ui.WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//li[@class='yui-ac-highlight']")))
    supplier_input.send_keys(Keys.RETURN)

    fill_in_data = {
        "Description":(By.ID, "textAreaID_11"),
        "Non-Catalog #": (By.NAME, "NonCatCatalogNumber"),
        "Qty":(By.NAME, "NonCatQuantity"),
        "Price (USD)":(By.NAME, "NonCatUnitPrice")
    }
    total_rows = df["Non-Catalog #"].count()
    # print()
    # print("Total Rows:", total_rows)
    # print()
    for index, row in df.iterrows():

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
    messagebox.showerror("Network Issue", "You are either having network problems, or took to long to do something.")
    sys.exit()
