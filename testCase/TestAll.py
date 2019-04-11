import xlrd
import ddt
import os
import unittest
from time import sleep
from loginPage import LoginPage
from todoApplyPage import TodoApplyPage
from selenium import webdriver
import HTMLTestRunner

def get_test_data(path, sheetname):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name(sheetname)
    nrows = sheet.nrows
    data = []
    if nrows > 1:
        for r in range(1, nrows):
            user = {}
            user["name"] = []
            user["username"] = []
            for c in range(0, 14, 2):
                user["name"].append(sheet.cell_value(r, c))
                user["username"].append(sheet.cell_value(r, c+1))
            user["value1"] = sheet.cell_value(r, 14)
            user["value2"] = sheet.cell_value(r, 15)
            user["value3"] = sheet.cell_value(r, 16)
            user["value4"] = sheet.cell_value(r, 17)
            data.append(user)
        return data
    else:
        return 0

# path = os.path.join(os.path.dirname(os.getcwd()),"testData\\HRMS.xlsx")
path = "HRMS.xlsx"
while 1:
    print("1 请假")
    print("2 加班")
    print("3 签卡")
    print("4 出差（无需购买飞机票）")
    print("5 出差（需要购买飞机票）")
    print("6 离职")
    print("7 离职结算")
    print("8 转正")
    print("请输入正确的申请类型编号：")
    t = input()
    if t == "1":
        sheetname = "请假"
        break
    elif t == "2":
        sheetname = "加班"
        break
    elif t == "3":
        sheetname = "签卡"
        break
    elif t == "4":
        sheetname = "出差"
        break
    elif t == "5":
        sheetname = "出差（飞机）"
        break
    elif t == "6":
        sheetname = "离职"
        break
    elif t == "7":
        sheetname = "离职结算"
        break
    elif t == "8":
        sheetname = "转正"
        break
    else:
        print("申请类型编号不正确，请重新输入：")


@ddt.ddt
class TestAll(unittest.TestCase):
    def setUp(self):
        print("开始测试")
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
    @ddt.data(*get_test_data(path, sheetname))
    def test_all(self, user):

        name = user["name"]
        username = user["username"]
        value1 = user["value1"]
        value2 = user["value2"]
        value3 = user["value3"]
        value4 = user["value4"]

        login_test = LoginPage(self.driver)
        apply_test = TodoApplyPage(self.driver)
        flag = 0

        for n in range(0, 7):
            login_test.open()
            login_test.login(username[n])
            if not flag:
                apply_test.shadow_click()
                if t == "1":
                    apply_test.apply_leave(value2, value3, value1)
                elif t == "2":
                    apply_test.apply_overtime(value2, value3, value1)
                elif t == "3":
                    apply_test.apply_sign(value2, value1)
                elif t == "4":
                    apply_test.apply_business(value2, value3, value4, value1)
                elif t == "5":
                    apply_test.apply_business_need_plane(value2, value3, value4, value1)
                elif t == "6":
                    apply_test.apply_dismission(value2)
                elif t == "7":
                    apply_test.apply_dismissionCheckout(value1)
                elif t == "8":
                    apply_test.apply_full(value1, value2, value3)
                else:
                    print("找不到该类型编号")
                    return 0
                print("申请人：", username[n], name[n])
                apply_test.next_processer(name[n+1])
                apply_test.logout()
                flag += 1
                sleep(1)
            else:
                print("审批人" + str(n) + "：", username[n], name[n])
                if t == "6" or t == "7":
                    apply_test.dismission_process(value2)
                else:
                    apply_test.process()
                result = apply_test.is_need_next_process(name[n+1])
                self.assertTrue(result != -1, "测试不通过")
                if result:
                    return 0

    def tearDown(self):
        print("结束测试\n")
        sleep(2)
        self.driver.quit()

if __name__ == '__main__':
    unittest.main()
#     unittest.TestLoader().loadTestsFromTestCase('TestAll')
#     unittest.TextTestRunner.run()
# testsuite = unittest.TestSuite()
# #     # 将测试用例加入到测试容器中
# # testsuite.addTest(unittest.makeSuite(TestAll))
# # # #
# # test_dir = os.getcwd()   # r'E:\workspace\HRMS'
# # # testsuite = unittest.defaultTestLoader.discover(test_dir, pattern='test_all_new2.py')
# #     # unittest.TextTestRunner().run(testsuite)
# # resultFilePath = os.path.join(os.path.dirname(test_dir),"testResult\\result.html")
# # fp = open(resultFilePath, 'wb')
# # runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title=sheetname+"自动化测试报告",description="用例执行情况")
# # runner.run(testsuite)
# # fp.close()