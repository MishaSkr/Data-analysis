# %%
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from bs4 import BeautifulSoup as bs
import re as re
import time
import pandas as pd
from selenium.webdriver.support.ui import Select
import os
from selenium.webdriver.chrome.options import Options

import numpy as np
from openpyxl import load_workbook


USERNAME = 
PASSWORD = 

file_path = os.path.join("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management")
#file_name= os.path.split("Auto_Projects.csv")

os.remove("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Projects export.csv")

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": file_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(chrome_options=chrome_options)

driver.get("https://app.activecollab.com/129853/reports/projects")
time.sleep(3)
email=driver.find_element(By.ID,"email")
email.send_keys(USERNAME)
password=driver.find_element(By.ID,"password")
password.send_keys(PASSWORD)
time.sleep(3)
password.send_keys(Keys.RETURN)

time.sleep(20)

select = Select(driver.find_element(By.CSS_SELECTOR,"tr:nth-of-type(6) > td:nth-of-type(2) > select"))
select.select_by_visible_text('Open and Completed')

budget=driver.find_element(By.XPATH, '//*[@id="reports_assignments"]/div/div[2]/div/div[1]/form/div/table/tbody/tr[6]/td[2]/label/input').click()
#select_report=driver.find_element(By.CSS_SELECTOR,".menu_inner")

export_button = driver.find_element(By.CSS_SELECTOR,".button_group:nth-child(1) .btn").click()

time.sleep(15)

driver.close()

try:
    os.rename("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/export.csv","/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Projects export.csv")
except FileNotFoundError:
    print("No file to rename")


file_path = os.path.join("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management")
#file_name= os.path.split("Auto_Projects.csv")

os.remove("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Tasks export.csv")

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": file_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(chrome_options=chrome_options)

#driver = webdriver.Chrome(PATH)
driver.get("https://app.activecollab.com/129853/reports/assignments")
time.sleep(3)
email=driver.find_element(By.ID,"email")
email.send_keys(USERNAME)
password=driver.find_element(By.ID,"password")
password.send_keys(PASSWORD)
time.sleep(3)
password.send_keys(Keys.RETURN)

time.sleep(20)

select = Select(driver.find_element(By.CSS_SELECTOR,".reports_table_wrapper+ .reports_table_wrapper tr:nth-child(5) .ng-not-empty"))
select.select_by_visible_text('Open and Completed')

tracked_time=driver.find_element(By.XPATH, '//*[contains(concat( " ", @class, " " ), concat( " ", "ng-empty", " " ))]').click()
#select_report=driver.find_element(By.CSS_SELECTOR,".menu_inner")

export_button = driver.find_element(By.CSS_SELECTOR,".button_group:nth-child(1) .btn").click()

time.sleep(55)

driver.close()

try:
    os.rename("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/export.csv","/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Tasks export.csv")
except FileNotFoundError:
    print("No file to rename")

file_path = os.path.join("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management")
#file_name= os.path.split("Auto_Projects.csv")

os.remove("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Time Report export.csv")

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": file_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(chrome_options=chrome_options)

#driver = webdriver.Chrome(PATH)
driver.get("https://app.activecollab.com/129853/reports/time-records")
time.sleep(3)
email=driver.find_element(By.ID,"email")
email.send_keys(USERNAME)
password=driver.find_element(By.ID,"password")
password.send_keys(PASSWORD)
time.sleep(3)
password.send_keys(Keys.RETURN)

time.sleep(20)

export_button = driver.find_element(By.CSS_SELECTOR,".button_group:nth-child(1) .btn").click()

time.sleep(70)

driver.close()

try:
    os.rename("/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/export.csv","/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Time Report export.csv")
except FileNotFoundError:
    print("No file to rename")


#Import all sourse datasheets

time = pd.read_csv('/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Time Report export.csv')
projects= pd.read_csv('/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Projects export.csv')
tasks= pd.read_csv('/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Tasks export.csv')
expenses= pd.read_csv('/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Expence export for 2022.csv')

rates= pd.read_excel('/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Rates Sale.xlsx')

# %%
#Change the data type of Tasks columns

task_dates = ['Created On', 'Start On', 'Due On', 'Completed On']
tasks = tasks.convert_dtypes()
tasks[task_dates] = tasks[task_dates].apply(pd.to_datetime)

#Rename the columns
tasks.rename(columns={'Project':'Project Name'}, inplace=True)
tasks.rename(columns={'Name':'Parent Name'}, inplace=True)
tasks.rename(columns={'Task List':'Task List Name'}, inplace=True)

#Filter Tracked time to null

tasks = tasks[tasks['Tracked Time'].isnull()]

tasks_3_columns = tasks[['Project Name','Task List Name','Parent Name']]

#print(tasks.dtypes)


#Change the data type of Project columns

project_dates = ['Created On', 'Completed On']
projects = projects.convert_dtypes()
projects[project_dates] = projects[project_dates].apply(pd.to_datetime)


#Change the data type of Expenses columns

expenses_dates = ['Record Date']
expenses = expenses.convert_dtypes()
expenses[expenses_dates] = expenses[expenses_dates].apply(pd.to_datetime)

#Rename the columns
expenses.rename(columns={'Value':'Expense Task Total'}, inplace=True)

expenses_4_columns = expenses[['Task List Name', 'Project Name', 'Expense Task Total']]

#Group by projects, then by tasks
expenses_4_columns = expenses_4_columns.groupby(['Project Name', 'Task List Name']).agg({'Expense Task Total': 'sum'}).reset_index()

#Create a joint 
expenses_4_columns['Joint'] = expenses_4_columns['Project Name']+' | '+expenses_4_columns['Task List Name']

expenses_2_columns = expenses_4_columns.copy()

expenses_3_columns = expenses_4_columns[['Project Name', 'Task List Name', 'Joint']]

expenses_2_columns = expenses_4_columns[['Joint', 'Expense Task Total']]


#Change the data type of Time columns

time_dates = ['Record Date']
time = time.convert_dtypes()
time[time_dates] = time[time_dates].apply(pd.to_datetime)

time_tasks = pd.concat([time, tasks_3_columns],axis=0)
time_tasks_rates = pd.merge(left=time_tasks,right=rates,how='left',left_on='Group Name',right_on='Group Name')
time_tasks_rates['Value GBP']=time_tasks_rates['Value']*time_tasks_rates['Rate']

time_tasks_rates['Budget Task Total'] = time_tasks_rates['Task List Name'].apply(lambda x: re.findall(r'\d+\.\d+|\d+,\d+\.\d+|\d+,\d+', x)[0] if re.findall(r'\d+\.\d+|\d+,\d+\.\d+|\d+,\d+', x) else None)
#time_tasks_rates['Budget Task Total'] = time_tasks_rates['Task List Name'].apply(lambda x: re.findall(r'\d+\.\d+|\d+,\d+', x)[0] if re.findall(r'\d+\.\d+|\d+,\d+', x) else None)
time_tasks_rates['Budget Task Total'] = time_tasks_rates['Budget Task Total'].str.replace(',','')

time_tasks_rates['Budget Task Total'] = pd.to_numeric(time_tasks_rates['Budget Task Total'], errors='coerce')
#time_tasks_rates['Budget Task Total']=float(time_tasks_rates['Budget Task Total'][0])


time_tasks_rates['Joint'] = time_tasks_rates['Project Name']+' | '+time_tasks_rates['Task List Name']

#______________________________________

base_for_aAntiRightJoin = time_tasks_rates.copy()

base_for_aAntiRightJoin = pd.merge(left=time_tasks_rates,right=expenses_4_columns,how='outer', indicator = True)

expenses_AntiRightJoin = base_for_aAntiRightJoin[(base_for_aAntiRightJoin._merge=='right_only')].drop('_merge', axis=1)

#time_r_e = time_r_e[time_r_e.isnull().any(axis=1)]

#join expenses
#time_r_e = pd.merge(left=time_with_rates,right=expenses_final,how='left',left_on='Joint',right_on='Joint')

expenses_AntiRightJoin = expenses_AntiRightJoin[['Project Name','Task List Name','Joint']]
time_tasks_rates_expenses = pd.concat([time_tasks_rates, expenses_AntiRightJoin],axis=0)
time_tasks_rates_expenses = pd.merge(left = time_tasks_rates_expenses, right = expenses_2_columns, how='left', left_on = 'Joint', right_on = 'Joint')
time_tasks_rates_expenses['Count'] = time_tasks_rates_expenses.groupby('Joint')['Joint'].transform('count')
time_tasks_rates_expenses['Budget Task Total'] = time_tasks_rates_expenses['Budget Task Total'].astype(float)
time_tasks_rates_expenses['Budget Task for Pivot'] = time_tasks_rates_expenses['Budget Task Total'] / time_tasks_rates_expenses['Count']
time_tasks_rates_expenses['Expense Task for Pivot'] = time_tasks_rates_expenses['Expense Task Total']/ time_tasks_rates_expenses['Count']

# Budget_Tracker = load_workbook('/Users/mikhailskrebnev/Library/CloudStorage/OneDrive-CoterieMarketing/Budget management/Python test output.xlsx')
# writer = pd.ExcelWriter('/Users/mikhailskrebnev/Library/CloudStorage/OneDrive-CoterieMarketing/Budget management/Python test output.xlsx',engine='openpyxl', mode='a', if_sheet_exists='replace')
# writer.Budget_Tracker=Budget_Tracker
# time_tasks_rates_expenses.to_excel(writer,sheet_name='Sheet2',index=False)
# writer.close()


output = time_tasks_rates_expenses
output.to_excel('/Users/mikhailskrebnev/Library/CloudStorage/OneDrive/Budget management/Sourse for budget tracker.xlsx', index=False)

