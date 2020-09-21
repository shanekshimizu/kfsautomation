from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome()


driver.get("https://kfs.hawaii.edu/kfs-prd6/portal.jsp")
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="iframe_portlet_container_table"]/tbody/tr[2]/td[2]/div[2]/div[2]/div/ul[5]/li[6]/a'))).click()

WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/div[1]/span/input'))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/div[1]/span/input'))).send_keys('foo')
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/div[2]/input'))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/div[2]/input'))).send_keys('foo')
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[1]/div[2]/form/p[2]/input[3]'))).click()

WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME,'iframe')))

WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="login-form"]/div[2]/div/label/input'))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="auth_methods"]/fieldset/div[1]/button'))).click()
driver.switch_to.default_content()

time.sleep(5)
driver.switch_to.default_content()

WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'/html/body/div[8]/div/iframe')))
iframethis2 = driver.find_element_by_id('iframeportlet')
driver.switch_to.frame(iframethis2)

WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/form/table/tbody/tr/td[2]/div/div[1]/div[2]/table[1]/tbody/tr[1]/td[2]/textarea'))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/form/table/tbody/tr/td[2]/div/div[1]/div[2]/table[1]/tbody/tr[1]/td[2]/textarea'))).send_keys('test text')
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/form/table/tbody/tr/td[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[2]/a/img'))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/form/table/tbody/tr/td[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[2]/div/input[1]'))).send_keys('/Users/shaneshimizu/Desktop/VSCode/automationtesting/Template_Fuel_Charges_cow_REV.csv')
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/form/table/tbody/tr/td[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[2]/div/input[2]'))).click()


driver.switch_to.default_content()

