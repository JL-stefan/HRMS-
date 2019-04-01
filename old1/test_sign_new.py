import xlrd
import ddt
import unittest
from time import sleep
from selenium import webdriver
from loginPage import LoginPage
from todoApplyPage import TodoApplyPage
import HTMLTestRunner

def get_test_data(path, sheetname):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name(sheetname)
    nrows = sheet.nrows
    data = []
    if nrows > 1:
        for r in range(1, nrows):
            user = {}
            user['username'] = sheet.cell_value(r, 0)
            user['password'] = sheet.cell_value(r, 1)
            user['date'] = sheet.cell_value(r, 2)
            user['reason'] = sheet.cell_value(r, 3)
            user['processer_name'] = []
            user['processer_username'] = []
            for c in range (4, 15, 2):
                user['processer_name'].append(sheet.cell_value(r, c))
                user['processer_username'].append(sheet.cell_value(r, c+1))
            data.append(user)
        return data
    else:
        return 0

# path = os.path.join(os.path.dirname(os.getcwd()), 'HRMS.xlsx')
path = "HRMS.xlsx"
sheetname = '签卡'

@ddt.ddt
class TestSign(unittest.TestCase):

    def setUp(self):
        print("开始测试")
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()

    @ddt.data(*get_test_data(path, sheetname))
    def test_sign(self,user):
        # 把已读取的数据信息赋值给变量
        username = str(user['username'])
        password = str(user['password'])
        sign_date = str(user['date'])
        reason = str(user['reason'])
        processer_name = user['processer_name']
        processer_username = user['processer_username']

        # 初始化
        login_test = LoginPage(self.driver)
        apply_test = TodoApplyPage(self.driver)
        # 标志位，判断走申请流程还是审批流程
        flag = 0
        for n in range(0, 6):
            login_test.open()
            # 申请流程
            if not flag:
                login_test.login(username, password)
                apply_test.shadow_click()
                apply_test.apply_sign(sign_date, reason)
                apply_test.next_processer(processer_name[n])
                apply_test.logout()
                flag += 1
                sleep(1)
            # 审批流程
            else:
                login_test.login(processer_username[n], password)
                apply_test.process()
                # 判断是否需要下一审批（是否有下一审批选择界面）
                result = apply_test.is_need_next_process(processer_name[n + 1])
                self.assertTrue(result != -1, "测试不通过")
                if result:
                    return 0

    def tearDown(self):
        print("结束测试\n")
        sleep(2)
        self.driver.quit()

if __name__ == '__main__':
    unittest.main()
#     # unittest.TestLoader().loadTestsFromTestCase('TestSign')
#     # unittest.TextTestRunner.run()
# #
# test_dir = r'E:\workspace\HRMS'
# suite = unittest.defaultTestLoader.discover(test_dir, pattern='test_sign_new.py')
#     # unittest.TextTestRunner().run(suite)
# fp = open('result.html', 'wb')
# runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title="test_result")
# runner.run(suite)