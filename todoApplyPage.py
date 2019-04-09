import os
from time import sleep
from basePage import BasePage
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select


class TodoApplyPage(BasePage):

    # 遮罩层元素
    shadow = (By.XPATH, '/html/body/div[2]/div')
    # 申请元素
    apply_btn = (By.XPATH, '//*[@id="app"]/div/div[1]/nav/div[1]/div')


    # 请假流程元素
    leave_btn = (By.XPATH, '//img[@title="请假"]')
    start_time = (By.XPATH, '//input[@placeholder="开始时间"]')
    end_time = (By.XPATH, '//input[@placeholder="结束时间"]')
    leave_reason = (By.XPATH, '//textarea[@placeholder="请输入请假原因"]')
    leave_commit = (By.XPATH, '//button/span[text()="提 交 "]')


    # 加班流程元素,开始时间,结束时间跟请假流程元素一样
    overtime_btn = (By.XPATH, '//img[@title="加班"]')
    close_overtime_date = (By.XPATH, '//label[text()="加班时长"]')
    choose_hours = (By.XPATH, '//input[@placeholder="请选择"]/parent::div/span/span/i')
    overtime_reason = (By.XPATH, '//textarea[@placeholder="请输入加班工作内容"]')
    overtime_commit = (By.XPATH, '//button/span[text()="提 交 "]')

    # 签卡流程元素
    sign_btn = (By.XPATH, '//img[@title="签卡"]')
    sign_date = (By.XPATH, '//input[@placeholder="请选择未刷卡时间"]')
    sure_date_btn = (By.XPATH, '/html/body/div[2]/div[2]/button[2]')
    sign_reason = (By.XPATH, '//textarea[@placeholder="请说明未刷卡的原因"]')
    commit_btn = (By.XPATH, '//button/span[text()="提 交 "]')

    # 出差流程元素,开始时间,结束时间跟请假流程元素一样
    business_btn = (By.XPATH, '//img[@title="出差"]')
    close_business_date = (By.XPATH, '//label[text()="出差地点"]')
    business_place= (By.XPATH, '//input[@placeholder="请输入出差地点"]')
    need_plane = (By.XPATH, '//span[text()="需行政人员购买机票"]')
    business_reason = (By.XPATH, '//textarea[@placeholder="请说明出差拟办事项"]')
    upload_business_file = (By.XPATH, '//*[@id="app"]/div/div[2]/div/div/div/div[3]/form/div[4]/div/div[1]/div[1]/input')
    business_commit = (By.XPATH, '//button/span[text()="提 交 "]')

    # 转正流程元素
    full_btn = (By.XPATH, '//img[@title="转正"]')
    work_content = (By.XPATH, '//textarea[@placeholder="请描述在职期间主要工作内容"]')
    change_content = (By.XPATH, '//textarea[@placeholder="请描述日常工作中存在的不足及改善措施"]')
    target_content = (By.XPATH, '//textarea[@placeholder="个人希望在公司的发展趋势及工作目标"]')
    full_commit = (By.XPATH, '//button/span[text()="提交"]')

    # 离职流程元素
    dismission_btn = (By.XPATH, '//img[@title="离职"]')
    company_reason = (By.XPATH, '//div[@class="dimission-panel__reason"]/div[2]/div[2]/div[4]')
    self_reason = (By.XPATH, '//div[@class="dimission-panel__reason"]/div[3]/div[2]/div[2]')
    expected_time = (By.XPATH, '//input[@placeholder="期望离职时间"]')
    dismission_commit = (By.XPATH, '//button/span[text()="提交"]')
    estimated_time = (By.XPATH, '//input[@placeholder="预计离职日期"]')

    # 离职结算流程元素
    dismissionCheckout_btn = (By.XPATH, '//img[@title="离职结算"]')
    work_handover = (By.XPATH, '//textarea[@placeholder="请说明工作交接情况"]')
    dismissionCheckout_commit = (By.XPATH, '//button/span[text()="提交 "]')
    real_time = (By.XPATH, '//input[@placeholder="实际离职时间"]')


    # 处理流程元素
    process_btn = (By.XPATH, '//table/tbody/tr[1]/td[6]/div/span')
    agree_btn = (By.XPATH, '//div[@class="action-group"]/button')
    ok_btn = (By.XPATH, '//footer/div/button[1]')

    # 处理时和处理后url
    process_before_url=''
    process_after_url=''

    def __init__(self, driver, timeout=20, url='https://hr-test.xiaojiaoyu100.com/flow/todoApply'):
        super().__init__(driver, timeout, url)

    # 处理遮罩层
    def shadow_click(self):
        self.find_element(*self.shadow).click()

    # 滚动界面到底部
    def move_to_foot(self):
        js = "var q=document.documentElement.scrollTop=100000"
        self.driver.execute_script(js)
        sleep(1)

    # 申请请假流程
    def apply_leave(self, start_time, end_time, reason):
        self.find_element(*self.apply_btn).click()
        self.find_element(*self.leave_btn).click()
        self.find_element(*self.start_time).send_keys(start_time)
        self.find_element(*self.end_time).send_keys(end_time)
        self.find_element(*self.leave_reason).send_keys(reason)
        self.move_to_foot()
        self.find_element(*self.leave_commit).click()

    #  申请加班流程
    def apply_overtime(self, start_time, end_time, reason):
        self.find_element(*self.apply_btn).click()
        self.find_element(*self.overtime_btn).click()
        self.find_element(*self.start_time).send_keys(start_time)
        self.find_element(*self.end_time).send_keys(end_time)
        self.find_element(*self.close_overtime_date).click()
        # js语句选择加班时数，这里默认选择3小时
        js = "document.getElementsByClassName('el-select-dropdown__item')[2].click()"
        self.driver.execute_script(js)
        self.find_element(*self.overtime_reason).send_keys(reason)
        self.move_to_foot()
        self.find_element(*self.overtime_commit).click()


    # 申请签卡流程
    def apply_sign(self, date, reason ):

        self.find_element(*self.apply_btn).click()
        self.find_element(*self.sign_btn).click()
        self.find_element(*self.sign_date).send_keys(date)
        self.find_element(*self.sure_date_btn).click()
        self.move_to_foot()
        self.find_element(*self.sign_reason).click()
        self.find_element(*self.sign_reason).clear()
        self.find_element(*self.sign_reason).send_keys(reason)
        sleep(1)
        self.find_element(*self.commit_btn).click()
        sleep(1)

    # 申请出差流程，不用乘坐飞机
    def apply_business(self, start_time, end_time, business_place, reason ):
        self.find_element(*self.apply_btn).click()
        self.find_element(*self.business_btn).click()
        self.find_element(*self.start_time).send_keys(start_time)
        sleep(1)
        self.find_element(*self.end_time).send_keys(end_time)
        sleep(1)
        self.find_element(*self.close_business_date).click()
        self.find_element(*self.business_place).send_keys(business_place)
        self.move_to_foot()
        self.find_element(*self.business_reason).send_keys(reason)
        self.find_element(*self.business_commit).click()

    # 申请出差流程，需要乘坐飞机
    def apply_business_need_plane(self, start_time, end_time, business_place, reason ):
        self.find_element(*self.apply_btn).click()
        self.find_element(*self.business_btn).click()
        self.find_element(*self.start_time).send_keys(start_time)
        sleep(1)
        self.find_element(*self.end_time).send_keys(end_time)
        sleep(1)
        self.find_element(*self.close_business_date).click()
        self.find_element(*self.business_place).send_keys(business_place)
        self.find_element(*self.need_plane).click()
        self.move_to_foot()
        self.find_element(*self.business_reason).send_keys(reason)
        file_path = os.path.join(os.path.dirname(os.getcwd()), 'testData\\《晓教育集团集中订购信息收集表》.xlsx')
        self.find_element(*self.upload_business_file).send_keys(file_path)
        # print(file_path)
        sleep(2)
        self.find_element(*self.business_commit).click()


    # 离职申请流程
    def apply_dismission(self, dismission_time):
        self.find_element(*self.apply_btn).click()
        self.find_element(*self.dismission_btn).click()
        self.find_element(*self.company_reason).click()
        self.find_element(*self.self_reason).click()
        self.find_element(*self.expected_time).send_keys(dismission_time)
        self.find_element(*self.dismission_commit).click()

    # 离职结算申请流程
    def apply_dismissionCheckout(self, reason):
        self.find_element(*self.apply_btn).click()
        self.find_element(*self.dismissionCheckout_btn).click()
        self.find_element(*self.work_handover).send_keys(reason)
        self.find_element(*self.dismissionCheckout_commit).click()

    # 选择下一个审核人流程
    def next_processer(self, name):
        # 定位审核人userid所在的元素input的祖父元素labek
        # nextPath =  '//input[@value=\''+str(userid)+'\']/parent::*/parent::label'

        # 通过文本内容定位,选择下一任审批人
        nextPath = '//span[text()=\'' + str(name) + '\']'
        processer = (By.XPATH, nextPath)
        # self.find_element(processer).click()
        processer_ele = WebDriverWait(self.driver, 20, 0.5).until(EC.presence_of_element_located(processer))
        processer_ele.click()
        sleep(1)
        self.find_element(*self.ok_btn).click()

    # 登出流程
    def logout(self):
        js = "document.getElementsByClassName('logout')[0].click()"
        self.driver.execute_script(js)

    # 审核人处理流程
    def process(self):
        global process_before_url
        global process_after_url
        self.find_element(*self.process_btn).click()
        self.move_to_foot()
        process_before_url = self.driver.current_url
        print("审核时url:%s" % process_before_url)
        self.find_element(*self.agree_btn).click()
        sleep(1)
        process_after_url = self.driver.current_url
        print("同意后url:%s" % process_after_url)

    # 离职流程审核人处理流程
    def dismission_process(self, dismission_time):
        global process_before_url
        global process_after_url
        self.find_element(*self.process_btn).click()
        self.move_to_foot()
        try:
            # 判断页面是否有预计离职时间元素
            # estimated_time = self.find_element(*self.estimated_time)
            estimated_time = self.driver.find_element(By.XPATH, '//input[@placeholder="预计离职日期"]')
        except:
            pass
        else:
            estimated_time.send_keys(dismission_time)
        try:
            # 判断页面是否有实际离职时间元素
            # real_time = self.find_element(*self.real_time)
            real_time = self.driver.find_element(By.XPATH, '//input[@placeholder="实际离职时间"]')
        except:
            pass
        else:
            real_time.send_keys(dismission_time)
        process_before_url = self.driver.current_url
        print("审核时url:%s"%process_before_url)
        self.find_element(*self.agree_btn).click()
        sleep(1)
        process_after_url = self.driver.current_url
        print("同意后url:%s" % process_after_url)

    def is_same_url(self):
        global process_before_url
        global process_after_url
        return process_before_url == process_after_url

    # 判断是否有下一审批人
    def is_next_processer_exist(self):
        try:
            self.find_element(*self.ok_btn)
            return True
        except:
            return False

    # # 针对请假/加班/签卡/出差判断是否需要下一审批（是否有下一审批选择界面）
    # def is_need_next_process(self, processer_name=''):
    #     if self.is_next_processer_exist():
    #         if processer_name:
    #             try:
    #                 self.next_processer(processer_name)
    #             except:
    #                 print("Failed：审批人"+processer_name+"找不到")
    #                 return -1
    #             self.logout()
    #             sleep(1)
    #         else:
    #             print("Failed：缺少下一审批人的数据")
    #             return -1
    #     elif processer_name:
    #         print("Warning：存在审批人"+processer_name+"数据冗余")
    #         return -1
    #     else:
    #         print("Passed：审批流程结束")
    #         return 1

    # 针对离职流程判断是否需要下一审批
    def is_need_next_process(self, processer_name=''):
        if self.is_same_url():
            if self.is_next_processer_exist():
                if processer_name:
                    try:
                        self.next_processer(processer_name)
                    except:
                        print("Failed：审批人"+processer_name+"找不到")
                        return -1
                    self.logout()
                    sleep(1)
                else:
                    print("Failed：缺少下一审批人的数据")
                    return -1
            else:
                print("Failed:页面存在信息未正确填写")
                return -1
        elif processer_name:
            print("Warning：存在审批人"+processer_name+"数据冗余")
            return -1
        else:
            print("Passed：审批流程结束")
            return 1

