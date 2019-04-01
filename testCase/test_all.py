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
            user["reason"] = sheet.cell_value(r, 14)
            user["start_time"] = sheet.cell_value(r, 15)
            user["end_time"] = sheet.cell_value(r, 16)
            user["place"] = sheet.cell_value(r, 17)
            data.append(user)
        return data
    else:
        return 0

path = os.path.join(os.path.dirname(os.getcwd()),"testData\\HRMS.xlsx")
while 1:
    print("1 请假")
    print("2 加班")
    print("3 签卡")
    print("4 出差（无需购买飞机票）")
    print("5 出差（需要购买飞机票）")
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
        reason = user["reason"]
        start_time = user["start_time"]
        end_time = user["end_time"]
        place = user["place"]

        login_test = LoginPage(self.driver)
        apply_test = TodoApplyPage(self.driver)
        flag = 0

        for n in range(0, 7):
            login_test.open()
            login_test.login(username[n])
            if not flag:
                apply_test.shadow_click()
                if t == "1":
                    apply_test.apply_leave(start_time, end_time, reason)
                elif t == "2":
                    apply_test.apply_overtime(start_time, end_time, reason)
                elif t == "3":
                    apply_test.apply_sign(start_time, reason)
                elif t == "4":
                    apply_test.apply_business(start_time, end_time, place, reason)
                elif t == "5":
                    apply_test.apply_business_need_plane(start_time, end_time, place, reason)
                else:
                    print("找不到该类型编号")
                    return 0
                print("申请人：", username[n], name[n])
                apply_test.next_processer(name[n+1])
                apply_test.logout()
                flag += 1
                sleep(1)
            else:
                apply_test.process()
                print("审批人" + str(n) + "：", username[n], name[n])
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
#     # unittest.TestLoader().loadTestsFromTestCase('TestAll')
#     # unittest.TextTestRunner.run()
# #
# test_dir = os.getcwd()   # r'E:\workspace\HRMS'
# suite = unittest.defaultTestLoader.discover(test_dir, pattern='test_all.py')
#     # unittest.TextTestRunner().run(suite)
# resultFilePath = os.path.join(os.path.dirname(os.getcwd()),"testResult\\result.html")
# fp = open(resultFilePath, 'wb')
# runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title="test_result")
# runner.run(suite)

