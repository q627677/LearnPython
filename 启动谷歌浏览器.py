
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

#设置谷歌浏览器地址  
__browser_url = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'  ##360浏览器的地址  
chrome_options = Options()
chrome_options.binary_location = __browser_url
browser = webdriver.Chrome(r'C:\Program Files (x86)\Google\Chrome\Application\chromedriver')  

# browser.get('http://www.baidu.com')
# browser.find_element_by_id("kw").send_keys("seleniumhq" + Keys.RETURN)  
# time.sleep(3)
# browser.quit()
browser.get(r'http://saishi.cnki.net/PaperIndex/ksf4e3ae35-1287-11ea-827b-801844e9f549')
