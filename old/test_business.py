import xlrd
import ddt
import unittest
from time import sleep
from loginPage import LoginPage
from todoApplyPage import TodoApplyPage
from selenium import webdriver

def get_test_data(path, sheetname):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name(sheetname)
    nrows = sheet.nrows
    data = []

    if nrows > 1:
        for i in range(1, nrows):
            user = {}
            user['username'] = sheet.cell_value(i, 0)
            user['password'] = sheet.cell_value(i, 1)
            user['start_time'] = sheet.cell_value(i, 2)
            user['end_time'] = sheet.cell_value(i, 3)
            user['place'] = sheet.cell_value(i, 4)
            user['reason'] = sheet.cell_value(i, 5)
            user['processer1_name'] = sheet.cell_value(i, 6)
            user['processer1_username'] = sheet.cell_value(i, 7)
            user['processer2_name'] = sheet.cell_value(i, 8)
            user['processer2_username'] = sheet.cell_value(i, 9)
            user['processer3_name'] = sheet.cell_value(i, 10)
            user['processer3_username'] = sheet.cell_value(i, 11)
            user['processer4_name'] = sheet.cell_value(i, 12)
            user['processer4_username'] = sheet.cell_value(i, 13)
            user['processer5_name'] = sheet.cell_value(i, 14)
            user['processer5_username'] = sheet.cell_value(i, 15)
            user['processer6_name'] = sheet.cell_value(i, 16)
            data.append(user)
        return data
    else:
        return 0

path = "HRMS.xlsx"
sheetname = '出差'

@ddt.ddt
class TestLeave(unittest.TestCase):

    def setUp(self):
        print("开始测试")
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()

    @ddt.data(*get_test_data(path, sheetname))
    def test_overtime(self, user):
        # 读取账号/密码/日期/原因等信息
        username = str(user['username'])
        password = str(user['password'])
        start_time = str(user['start_time'])
        end_time = str(user['end_time'])
        place = str(user['place'])
        reason = str(user['reason'])
        processer1_name = str(user['processer1_name'])
        processer1_username = user['processer1_username']
        processer2_name = user['processer2_name']
        processer2_username = user['processer2_username']
        processer3_name = user['processer3_name']
        processer3_username = user['processer3_username']
        processer4_name = user['processer4_name']
        processer4_username = user['processer4_username']
        processer5_name = user['processer5_name']
        processer5_username = user['processer5_username']
        processer6_name = user['processer6_name']



        # 登录流程
        login_test = LoginPage(self.driver)
        login_test.open()
        login_test.login(username, password)
        # sleep(1)
        # apply_test.logout()
        # sleep(2)
        # 开始申请出差
        apply_test = TodoApplyPage(self.driver)
        apply_test.shadow_click()
        apply_test.apply_business(start_time, end_time, place, reason)
        # 选择下一个审批人
        apply_test.next_processer(processer1_name)
        # 退出登录
        apply_test.logout()
        sleep(1)

        # 审批人1 登录同意操作流程
        login_test.open()
        login_test.login(processer1_username, password)
        apply_test.process()
        # 判断是否需要下一审批（是否有下一审批选择界面）
        result = apply_test.is_need_next_process(processer2_name)
        self.assertTrue(result != -1, "测试不通过")
        if result:
            return 0

        # 审批人2登录同意签卡操作流程
        login_test.open()
        login_test.login(processer2_username, password)
        apply_test.process()
        result = apply_test.is_need_next_process(processer3_name)
        self.assertTrue(result != -1, "测试不通过")
        if result:
            return 0

        # 审批人3登录同意签卡操作流程
        login_test.open()
        login_test.login(processer3_username, password)
        apply_test.process()
        result = apply_test.is_need_next_process(processer4_name)
        self.assertTrue(result != -1, "测试不通过")
        if result:
            return 0

        # 审批人4登录同意签卡操作流程
        login_test.open()
        login_test.login(processer4_username, password)
        apply_test.process()
        result = apply_test.is_need_next_process(processer5_name)
        self.assertTrue(result != -1, "测试不通过")
        if result:
            return 0

        # 审批人5登录同意签卡操作流程
        login_test.open()
        login_test.login(processer5_username, password)
        apply_test.process()
        result = apply_test.is_need_next_process(processer6_name)
        self.assertTrue(result != -1, "测试不通过")
        if result:
            return 0

        # if self.login_success_url in self.driver.current_url:
        #     print(msg, login_test.login_success())
        #     self.assertTrue(msg in login_test.login_success(), '帐号密码正确，测试不通过')
        # else:
        #     print(msg, login_test.message())
        #     self.assertTrue(msg in login_test.message(), '帐号或者密码错误时，测试不通过')

    def tearDown(self):
        print("结束测试\n")
        sleep(2)
        self.driver.quit()

if __name__ == '__main__':
    unittest.main()
    # unittest.TestLoader().loadTestsFromTestCase('TestLeave')
    # unittest.TextTestRunner.run()