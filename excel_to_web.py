# -*- encoding: utf-8 -*-
"""
@Author      : Alfred T
@Time        : 2020/5/16 22:17
@description : 
"""

import xlrd
from selenium import webdriver

# driver = webdriver.Ie()
# driver.get("https://www.baidu.com")

data = xlrd.open_workbook("C:\\Users\\ty\\Desktop\\test.xlsx")
print(data.sheet_names())
table = data.sheet_by_name('Sheet1')
rowNum = table.nrows
colNum = table.ncols
farst = table.row_values(0)
print(farst)
