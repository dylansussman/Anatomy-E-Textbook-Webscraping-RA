import os
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

DRIVER_PATH = '/usr/local/bin/chromedriver'
driver = webdriver.Chrome(executable_path=DRIVER_PATH)
driver.get('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu')

username: WebElement = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'username')))
password: WebElement = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'password')))
username.send_keys('sussman.47')
password.send_keys('Banbury3117&')
driver.find_element(By.ID, 'submit').click()
