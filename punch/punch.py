from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from pandas import read_excel
data=read_excel('data.xlsx',index_col=0,header=1)
driver = webdriver.Chrome('C:/Users/F128537908/chromedriver.exe')
driver.get('http://jshradbs01/attmanual/default.aspx')
Account=driver.find_element_by_name('tAccount')
Account.send_keys(data.loc[1,'帳號'])
PASSWORD=driver.find_element_by_name('tPASSWORD')
PASSWORD.send_keys(data.loc[1,'密碼'])
PASSWORD.send_keys(Keys.RETURN)