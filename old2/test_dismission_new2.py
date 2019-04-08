import xlrd
import ddt
import unittest
import os
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
            user['name'] = []
            user['username'] = []
            for c in range(0, 13, 2):
                user['name'].append(sheet.cell_value(r, c))
                user['username'].append(sheet.cell_value(r, c + 1))
            user['start_time'] = sheet.cell_value(r, 14)
            data.append(user)
        return data
    else:
        return 0

# path = os.path.join(os.path.dirname(os.getcwd()), 'HRMS.xlsx')
path = os.path.join(os.path.dirname(os.getcwd()), "testData\\HRMS.xlsx")
sheetname = '离职'

@ddt.ddt
class TestSign(unittest.TestCase):

    def setUp(self):
        print("开始测试")
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()

    @ddt.data(*get_test_data(path, sheetname))
    def test_sign(self,user):
        # 把已读取的数据信息赋值给变量
        name = user['name']
        username = user['username']
        start_time = str(user['start_time'])

        # 初始化
        login_test = LoginPage(self.driver)
        apply_test = TodoApplyPage(self.driver)
        # 标志位，判断走申请流程还是审批流程
        flag = 0
        for n in range(0, 7):
            login_test.open()
            login_test.login(username[n])
            # 申请流程
            if not flag:
                print("申请人：", username[n], name[n])
                apply_test.shadow_click()
                apply_test.apply_dismission(start_time)
                apply_test.next_processer(name[n + 1])
                apply_test.logout()
                flag += 1
                sleep(1)
                # 审批流程
            else:
                print("审批人" + str(n) + "：", username[n], name[n])
                apply_test.process()
                # 判断是否需要下一审批（是否有下一审批选择界面）
                result = apply_test.is_need_next_process(name[n + 1])
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
# # suite = unittest.defaultTestLoader.discover(test_dir, pattern='test_sign_new.py')
# #     # unittest.TextTestRunner().run(suite)
# # fp = open('result.html', 'wb')
# # runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title="test_result")
# # runner.run(suite)