from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome()


driver.get("https://gmail.hawaii.edu")
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/div[1]/span/input'))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/div[1]/span/input'))).send_keys('user')
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/div[2]/input'))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/div[2]/input'))).send_keys('pass')
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/p[2]/input[3]'))).click()

WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(., "Send Me a Push")]'))).click()


