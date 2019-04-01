from time import sleep
from basePage import BasePage
from selenium.webdriver.common.by import By

class LoginPage(BasePage):

    user_loc = (By.ID, 'signin-username')
    pwd_loc = (By.ID, 'signin-password')
    submit_loc = (By.ID, 'signin-btn')

    def __init__(self, driver, timeout=20, url='https://hr-test.xiaojiaoyu100.com/flow/todoApply'):
        super().__init__(driver, timeout, url)

    def open(self, url='https://hr-test.xiaojiaoyu100.com/flow/todoApply'):
        super().open(url)

    def login(self, username, password="2009xabc"):
        self.find_element(*self.user_loc).send_keys(username)
        self.find_element(*self.pwd_loc).send_keys(password)
        sleep(2)
        self.find_element(*self.submit_loc).click()
        sleep(2)
