#!/usr/bin/python3

import os, sys, re
import time, calendar
import xlrd, xlwt
from xlutils.copy import copy as xlcopy

def xtest_date(ui_year):
    list_week_map = ("一", "二", "三", "四", "五", "六", "日",)
    for month in range(1, 13):
        num_day = calendar.monthrange(ui_year, month)[1]
        for day in range(1, num_day + 1):
            ui_weekday = calendar.weekday(ui_year, month, day)
            str_date = "%d-%02d-%02d, 星期%s" % (ui_year, month, day, list_week_map[ui_weekday])
            print(str_date)

def xwrite_workdiary_excel(ui_year):
    str_sheetname = '%s' % (ui_year)
    str_bookname = './bin/work_diary_%s.xls' % (ui_year)
    ui_cellwidth = 1178 #2356 = std_cellwidth

    if ( os.access(str_bookname, os.F_OK) ):  # tydbg
        exist_workbook = xlrd.open_workbook(str_bookname)
        for exist_sheetname in exist_workbook.sheet_names():
            if (exist_sheetname == str_sheetname):
                print('Error! the input sheetname:%s is exist' % (str_sheetname))
                exit(-1)
        workbook = xlcopy(exist_workbook)
    else:
        workbook = xlwt.Workbook(encoding='ascii')

    worksheet = workbook.add_sheet(str_sheetname)
    alighment = xlwt.Alignment()
    alighment.horz = 2  # 0:通用， 1-3:左中右
    alighment.vert = 1  # 0-2：上中下
    alighment.wrap = 1  # 1：自动换行
    style = xlwt.XFStyle()
    style.alignment = alighment

    worksheet.write(0, 0, '日期'  , style)
    worksheet.write(0, 1, '星期'  , style)
    worksheet.write(0, 2, '内容'  , style)
    worksheet.write(0, 3, '关键字', style)
    worksheet.col(0).width = ui_cellwidth * 2
    worksheet.col(1).width = ui_cellwidth * 1
    worksheet.col(2).width = ui_cellwidth * 20
    worksheet.col(3).width = ui_cellwidth * 4

    list_color = [44, 47, 57]  # [浅蓝，淡黄，浅绿]
    list_week_map = ("一", "二", "三", "四", "五", "六", "日")
    ui_rowcnt = 1
    for month in range(1, 13):
        num_day = calendar.monthrange(ui_year, month)[1]
        for day in range(1, num_day + 1):
            ui_weekday = calendar.weekday(ui_year, month, day)

            pattern = xlwt.Pattern()
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = list_color[month%3]

            borders = xlwt.Borders()  # Create Borders
            borders.left   = xlwt.Borders.THIN
            borders.right  = xlwt.Borders.THIN
            borders.top    = xlwt.Borders.THIN
            borders.bottom = xlwt.Borders.THIN
            borders.left_colour   = 0x40
            borders.right_colour  = 0x40
            borders.top_colour    = 0x40
            borders.bottom_colour = 0x40

            alighment = xlwt.Alignment()
            alighment.vert = 1 #0-2：上中下
            alighment.wrap = 1 #1：自动换行

            style = xlwt.XFStyle()
            style.pattern = pattern
            style.borders = borders
            style.alignment = alighment

            worksheet.write(ui_rowcnt, 0, '%02d/%02d' % (month,day), style)
            worksheet.write(ui_rowcnt, 1, '%s' % (list_week_map[ui_weekday]), style)
            worksheet.write(ui_rowcnt, 2, '', style)
            worksheet.write(ui_rowcnt, 3, '', style)
            ui_rowcnt += 1

    workbook.save(str_bookname)

if __name__ == '__main__':
    ui_year = int(sys.argv[1])
    print(ui_year)

    xwrite_workdiary_excel(ui_year)

# end of main