#!/usr/bin/env python
#encoding:utf-8

import os
import shutil
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import re
import openpyxl


def txt_list(txt_name, data=[]):
    file = open(txt_name, 'r')
    for line in file.readlines():
        line = line.strip('\n')
        line = re.split("[ |]", line)
        line_ret = [str for str in line if str not in ['', '', None]]
        data.append(line_ret)
    file.close()
    return data


def data_xlsx(filename, xlsxname, data=[]):
    wb = openpyxl.load_workbook(filename)
    ws = wb.get_sheet_by_name(u'原始数据')
    for i in xrange(len(data)):
        nums = len(data[i])
        for j in xrange(nums):
            try:
                data[i][j] = float(data[i][j])
            except ValueError:
                pass
            finally:
                ws.cell(row=i + 1, column=j + 1, value=data[i][j])
    wb.save(xlsxname)


if __name__ == '__main__':

    crbt_path = os.path.abspath('./crbt_vpn')
    sms_path = os.path.abspath('./sms')
    back_path = os.path.abspath('./backup')
    try:
        crbt_filename = os.listdir(crbt_path)[0]
        sms_filename = os.listdir(sms_path)[0]
    except:
        print '无原始文件'
        sys.exit()

    crbt_list = []
    sms_list = []

    txt_list(crbt_path + os.path.sep + crbt_filename, crbt_list)
    txt_list(sms_path + os.path.sep + sms_filename, sms_list)

    date_crbt= crbt_filename.split('.')[2]
    date_sms = sms_list[215][0]
    crbt_xlsxname = u'智能网和彩铃用户统计'+str(date_crbt)+'.xlsx'
    sms_xlsxname = u'短号短（彩）信业务统计'+str(date_sms)+'.xlsx'

    data_xlsx(filename='crbt_vpn_tj.xlsx', xlsxname=crbt_xlsxname, data=crbt_list)
    data_xlsx(filename='sms_tj.xlsx', xlsxname=sms_xlsxname, data=sms_list)

    if os.path.exists(back_path + os.path.sep + crbt_filename):
        os.remove(crbt_path + os.path.sep + crbt_filename)
    else:
        shutil.move(crbt_path + os.path.sep + crbt_filename, back_path)

    if os.path.exists(back_path + os.path.sep + sms_filename):
        os.remove(sms_path + os.path.sep + sms_filename)
    else:
        shutil.move(sms_path + os.path.sep + sms_filename, back_path)

    #os.environ['date_crbt']=str(date_crbt)
    #os.environ['crbt_xlsxname']=str(crbt_xlsxname)
    #os.environ['sms_xlsxname']=str(sms_xlsxname)
    #tar_cmd='tar zcf  tj${date_crbt}.tar.gz $crbt_xlsxname $sms_xlsxname'
    #os.system(tar_cmd)