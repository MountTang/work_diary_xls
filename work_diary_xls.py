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
            borders.left   = xlwt.Borders.THIN # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
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


'''
# 时间戳：
```
ticks = time.time()
print ("当前时间戳为:", ticks)
% 当前时间戳为: 1459996086.7115328
```

# 获取当前时间
localtime = time.localtime(time.time())
print ("本地时间为 :", localtime)
%本地时间为 : time.struct_time(tm_year=2016, tm_mon=4, tm_mday=7, \
              tm_hour=10, tm_min=28, tm_sec=49, tm_wday=3, tm_yday=98, tm_isdst=0)
6 tm_wday 0到6 (0是周一)
7 tm_yday 一年中的第几天，1 到 366
8 tm_isdst 是否为夏令时，值有：1(夏令时)、0(不是夏令时)、-1(未知)，默认 -1

# 格式化日期，讲某一日期转换为时间戳

### 格式化成2016-03-20 11:45:39形式
print (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

### 将格式字符串转换为时间戳
print (time.mktime(time.strptime("2016-04-07 10:29:46","%Y-%m-%d %H:%M:%S")))

### 格式化符号：
%y 两位数的年份表示（00-99）
%Y 四位数的年份表示（000-9999）
%m 月份（01-12）
%d 月内中的一天（0-31）
%H 24小时制小时数（0-23）
%I 12小时制小时数（01-12）
%M 分钟数（00=59）
%S 秒（00-59）
%w 星期（0-6），星期天为星期的开始
%W 一年中的星期数（00-53）星期一为星期的开始

# time模块常用函数
* time.clock( )
用以浮点数计算的秒数返回当前的CPU时间。

* time.mktime(tupletime)
接受时间元组并返回时间戳（1970纪元后经过的浮点秒数）。

* time.sleep(secs)
推迟调用线程的运行

* time.strftime(fmt, tupletime)
接收以时间元组，并返回以可读字符串表示的当地时间，格式由fmt决定。

* time.strptime(str,fmt='%a %b %d %H:%M:%S %Y')  # Thu Apr 07 10:25:09 2016
* time.strptime(str,fmt='%Y-%m-%d %H:%M:%S')     # 2016-04-07 10:29:46
根据fmt的格式把一个时间字符串解析为时间元组。

* time.time( )
返回当前时间的时间戳（1970纪元后经过的浮点秒数）。


# 日历（Calendar）模块
* calendar.calendar(year,w=2,l=1,c=6)
返回一个多行字符串格式的year年年历，
3个月一行，间隔距离为c。\
每日宽度间隔为w字符。每行长度为21* W+18+2* C。\
l是每星期行数。

* calendar.month(year,month,w=2,l=1)
返回一个多行字符串格式的year年month月日历，两行标题，一周一行。\
每日宽度间隔为w字符。每行的长度为7* w+6。l是每星期的行数

* calendar.weekday(year,month,day)
返回给定日期的日期码。0（星期一）到6（星期日）。月份为 1（一月） 到 12（12月）。

* calendar.monthrange(year,month)
返回两个整数。第一个是该月的星期几的日期码，第二个是该月的日期码。\
日从0（星期一）到6（星期日）;月从1到12。

* calendar.setfirstweekday(weekday)
设置每周的起始日期码。0（星期一）到6（星期日）。

* calendar.monthcalendar(year,month)
返回一个整数的单层嵌套列表。每个子列表装载代表一个星期的整数。\
Year年month月外的日期都设为0;范围内的日子都由该月第几日表示，从1开始。

'''