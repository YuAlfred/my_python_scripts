# -*- encoding: utf-8 -*-
"""
@Author      : Alfred T
@Time        : 2021/10/31 12:14
@description : 
"""

import xlrd
import xlwt
import xlutils

# 打开某一excel
curExcel = xlrd.open_workbook("./成都市郫都区绵实外国语学校(十二年一贯制学校).xlsx")
# 打开某一sheet
curSheet = curExcel.sheet_by_name("教基1001_十二年一贯制学校")
# 读取某一单元格
value = curSheet.cell(2, 2)


