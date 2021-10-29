# -*- encoding: utf-8 -*-
"""
@Author      : Alfred T
@Time        : 2020/5/16 22:17
@description : 
"""

import xlrd
from selenium import webdriver
import time


# web初始化
def web_init():
    # # 打开ie浏览器
    # driver = webdriver.Ie()
    # 打开Chrome浏览器
    options = webdriver.ChromeOptions()
    options.add_argument('lang=zh_CN.UTF-8')
    driver = webdriver.Chrome(chrome_options=options)

    driver.get("http://usp.cdrcbank.com/usp/")
    # driver.get("https://www.baidu.com/")

    # 等待输入后开始
    print("=====前置步骤结束后按下回车后开始程序=====")

    input()

    print("=====开始执行=====")
    # 获取打开的多个窗口句柄
    windows = driver.window_handles
    # 切换到当前最新打开的窗口
    driver.switch_to.window(windows[-1])
    # 切换到iframe
    driver.switch_to.frame("main-content-container")

    return driver


# 执行一次新增操作
# driver:浏览器驱动
def new_costumer(driver, name):
    # 点击新增按钮
    driver.find_element_by_xpath("//*[@id='global-query-table-div']/div/div[1]/div/div/a[3]/span/span/span[1]").click()
    # /html/body/div[2]/div/div[1]/div/div/a[3]/span/span/span[1]
    # /html/body/div[2] == //*[@id='global-query-table-div']
    time.sleep(1)

    # 第一个个输入框
    # 下拉按钮 orderBusiType 1198
    driver.find_element_by_xpath(
        "//*[@id='pms.service.booking.bookingCreateWindow']/div[2]/div/div[1]/span/div/table[1]/tbody/tr/td[2]/div/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[2]/div").click()
    #    /html/body/div[9]/div[2]/div == //*[@id="ext-comp-1075"]
    #    //*[@id='pms.service.booking.bookingCreateWindow'] == /html/body/div[6]
    #    /html/body/div[6]/div[2]/div/div[1]/span/div/table[1]/tbody/tr/td[2]/div/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[2]/div

    # 选择第二个
    driver.find_element_by_xpath("//*[@role = 'option' and text()='个人存款']").click()
    # /html/body/div[10]

    # 第二个输入框
    driver.find_element_by_xpath("//*[@name = 'cusName']").send_keys(name)

    # 第三个输入框（时间）ext-gen1195
    # 时间按钮
    driver.find_element_by_xpath(
        "//*[@id='pms.service.booking.bookingCreateWindow']/div[2]/div/div[1]/span/div/table[2]/tbody/tr/td[2]/div/div/div/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div").click()

    # 选择今天 button-1064-btnIconEl
    driver.find_element_by_xpath("//*[contains(@class,'x-datepicker')]/div/div[2]/a/span/span/span[2]").click()
    #  /html/body/div[11]/div/div[2]/a/span/span/span[2]
    #  /html/body/div[11] == //*[@class= 'x-datepicker']

    # 第四个输入框（金额） textfield-1050-inputEl
    time.sleep(0.5)
    driver.find_element_by_xpath("//*[@name = 'orderAmt']").send_keys("100")

    # 机构选择按钮 button-1054-btnIconEl
    driver.find_element_by_xpath(
        "//*[@id='pms.service.booking.bookingCreateWindow']/div[2]/div/div[1]/span/div/table[4]/tbody/tr/td[2]/div/div/div/a/span/span/span[2]").click()
    # /html/body/div[8]/div[2]/div/div[1]/span/div/table[4]/tbody/tr/td[2]/div/div/div/a/span/span/span[2]

    # 一级下拉菜单展开  //*[@id="each-busi-org-select-window"] == /html/body/div[10]
    driver.find_element_by_xpath(
        "//*[@id='each-busi-org-select-window']/div[2]/div/div[2]/div/table/tbody/tr/td[1]/div/img[1]").click()
    # 睡眠两秒避免系统反应不及时
    time.sleep(1.5)

    # 金牛支行 ext-gen1303 /html/body/div[10]/div[2]/div/div[2]/div/table/tbody/tr[14]/td[1]/div/span //*[@id="each-busi-org-select-window"]
    # /html/body/div[11]/div[2]/div/div[2]/div/table/tbody/tr[17]/td[1]/div/span
    # driver.find_element_by_xpath(
    #     "//*[@id='each-busi-org-select-window']/div[2]/div/div[2]/div/table/tbody/tr[14]/td[1]/div/img[2]").click()
    driver.find_element_by_xpath('//*[contains(text(), "金牛支行管理中心")]/../img[2]').click()
    time.sleep(2.5)

    # 金牛成化支行 ext-gen1373
    # /html/body/div[11]/div[2]/div/div[2]/div/table/tbody/tr[25]/td[1]/div/span
    # driver.find_element_by_xpath(
    #     "//*[@id='each-busi-org-select-window']/div[2]/div/div[2]/div/table/tbody/tr[23]/td[1]/div/span").click()
    try:
        bank_name = driver.find_element_by_xpath('//span[contains(text(), "金牛化成支行")]/..')
        bank_name.click()
    except Exception as e:
        print("支行名字没取到")
    time.sleep(1.5)

    # 确认 button-1071-btnIconEl
    driver.find_element_by_xpath(
        "//*[@id='each-busi-org-select-window']/div[2]/div/div[3]/div/div/a[1]/span/span/span[2]").click()
    # /html/body/div[12]/div[2]/div/div[3]/div/div/a[1]/span/span/span[2]
    time.sleep(1.5)

    # 保存按钮 button-1056-btnIconEl
    driver.find_element_by_xpath(
        "//*[@id='pms.service.booking.bookingCreateWindow']/div[2]/div/div[2]/div/div/a[1]/span/span/span[2]").click()
    # /html/body/div[6]/div[2]/div/div[2]/div/div/a[1]/span/span/span[2]

    time.sleep(1)

    # 操作成功的确定 button-1005-btnIconEl
    driver.find_element_by_xpath("//*[contains(@class,'x-message-box')]/div[3]/div/div/a[1]/span/span/span[2]").click()
    # /html/body/div[8]/div[3]/div/div/a[1]/span/span/span[2]
    # /html/body/div[8] = //*[contains(@class,'x-message-box')]
    time.sleep(0.5)


if __name__ == "__main__":
    data = xlrd.open_workbook("./格式模板 - 副本.xlsx")
    # 选择Sheet2
    table = data.sheet_by_name('Sheet2')
    rowNum = table.nrows
    # web初始化
    driver = web_init()
    # 每一行入录一次
    for i in range(rowNum):
        name = table.row_values(i)[0]
        print(str("==================正在写入：" + name + "=================="))
        new_costumer(driver, name)
