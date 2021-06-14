import pandas as pd
import os.path as osp
import xlrd
import xlwt
import math
import datetime
import time

def timer(func):
    def wrapper():
        start=time.time()
        func()
        end=time.time()
        return end-start
    return wrapper

def date_convert(org_date):
    if type(org_date)== float:
        delta = pd.Timedelta(str(org_date)+'D')
        real_time = pd.to_datetime('1899-12-30')+delta
        #tmp = str(real_time).split(' ')[0]
        tmp = real_time
    elif type(org_date) == str:
        tmp = pd.to_datetime(org_date)
    else:
        tmp = 0
    return tmp


# 判断日期是否为当月最后一天
def day_vote(date):
    month_29 = [2000, 2004, 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040, 2044]
    month_30 = [4, 6, 9, 11]
    offset = 0
    tmp = date.strftime('%Y-%m-%d').split('-')
    year = int(tmp[0])
    month = int(tmp[1])
    day = int(tmp[2])
    if day < 28:
        pass
    elif day == 28:
        if not month == 2:
            pass
        else:
            if year not in month_29:
                offset = 1
            else:
                pass
    elif day == 29:
        if month == 2:
            offset = 1
        else:
            pass
    elif day == 30 and (month in month_30):
        offset = 1
    elif day == 31:
        offset = 1
    return offset


# 月份迭代

def iter_for_month(init_date, fix_num):
    year, month, day = init_date.strftime('%Y-%m-%d').split('-')
    year, month, day = int(year), int(month), int(day)
    month2 = month + fix_num
    year_offset = 0
    if month2 <= 12:
        pass
    elif month2 <= 24:
        year_offset = 1
        month2 = month2 - 12
    else:
        year_offset, month2 = divmod(month2, 12)
        if month2 == 0:
            month2 = 12
            year_offset = year_offset - 1
    if len(str(month2)) < 2:
        month2 = '0' + str(month2)
    year, month2 = str(year + year_offset), str(month2)
    return year + month2


# 初始化生成文件
def build_init_xlbook(org_data, save_path):
    # init the rebuild file
    rebuild_file = xlwt.Workbook(encoding='utf-8')
    table = rebuild_file.add_sheet('1')
    table.write(0, 0, '站址名称')
    table.write(0, 1, '站址编码')
    table.write(0, 2, '场租合同期总金额')
    table.write(0, 3, '支付账期')
    table.write(0, 4, '账期应付金额')
    table.write(0, 5, 'key')

    # init the var
    timeline = 0
    timeline_history = [0]
    data_nums = org_data.nrows
    anchor = 0
    for i in range(1, data_nums):
        now = datetime.datetime.now()
        state_name = org_data.cell(i, 0).value
        state_id = org_data.cell(i, 1).value
        date0 = date_convert(org_data.cell(i, 2).value)  # start_date Type->TimeStamp
        date1 = date_convert(org_data.cell(i, 3).value)  # end_date
        fee = org_data.cell(i, 4).value
        offset = day_vote(date1)
        timeline = (date1.year - date0.year) * 12 + (date1.month - date0.month) + offset
        if timeline < 1:
            timeline = 1
        timeline_history.append(timeline)
        month_fee = fee / timeline
        # month_fee=fee
        anchor = anchor + timeline_history[i - 1]
        print('当前处理第{}个数据，锚点为{}行，起始日期为{}，终止日期为{}，账期为{}'.format(i, anchor, date0, date1, timeline))
        month_fee = round(fee / timeline, 2)
        for j in range(timeline):
            table.write(j + anchor + 1, 0, state_name)
            table.write(j + anchor + 1, 1, state_id)
            table.write(j + anchor + 1, 2, fee)
            table.write(j + anchor + 1, 3, iter_for_month(date0, j))
            table.write(j + anchor + 1, 4, month_fee)
            table.write(j + anchor + 1, 5, state_id + '-' + iter_for_month(date0, j))
        print('第{}个数据已处理完成，耗时{}'.format(i, datetime.datetime.now() - now))
    rebuild_file.save(save_path)
    return None