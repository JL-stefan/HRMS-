from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

class BasePage():

    def __init__(self, driver, timeout, url):
        self.driver = driver
        self.timeout = timeout
        self.url = url

    def open(self, url):    #, cookies
        self.driver.get(url)
        # for cookie in cookies:
        #     self.driver.add_cookie(cookie_dict=cookie)
        #     print(cookie)


    # def find_element(self, *loc):
    #     return WebDriverWait(self.driver, 6, 0.5).until(
    #         EC.presence_of_element_located(loc)
    #     )
    #
    def find_element(self, *loc):
        try:
            ele = WebDriverWait(self.driver, 6, 0.5).until(EC.presence_of_element_located(loc))
        except:
            print("找不到该元素："+loc)
            return -1
        else:
            return ele