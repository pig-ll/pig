# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
import os
import xlrd
from datetime import date,datetime
from xlrd import xldate_as_tuple
import json
import re
import pymysql

conn = pymysql.connect(host="47.104.11.14", user="root", passwd="mysqldb", db="dp_apt",charset='utf8')
cursor = conn.cursor()
def get_merged_cells(sheet):
    # [(4, 5, 2, 4), (5, 6, 2, 4), (1, 4, 3, 4)]
    # (4, 5, 2, 4) 的含义为：行 从下标4开始，到下标5（不包含）  列 从下标2开始，到下标4（不包含），为合并单元格
    return sheet.merged_cells

def get_merged_cells_value(sheet, row_index, col_index):
    # 先判断给定的单元格，是否属于合并单元格；
    # 如果是合并单元格，就返回合并单元格的内容
    merged = get_merged_cells(sheet)
    for (rlow, rhigh, clow, chigh) in merged:
        if (row_index >= rlow and row_index < rhigh):
            if (col_index >= clow and col_index < chigh):
                cell_value = sheet.cell_value(rlow, clow)
                # print('该单元格[%d,%d]属于合并单元格，值为[%s]' % (row_index, col_index, cell_value))
                return cell_value,(rlow, rhigh, clow, chigh)
                # break
    return (sheet.cell_value(row_index,col_index),(0,row_index+1,0,0))

def read_excel(rootDir):
    for root,dirs,files in os.walk(rootDir):
        for file in files:
            file_name = os.path.join(root,file)
            (filepath, tempfilename) = os.path.split(file_name)
            (filename, extension) = os.path.splitext(tempfilename)
            print filename,file_name
            path_key = filename
            path_value = file_name
            inpath = path_value
            uipath = file_name
            ExcelFile=xlrd.open_workbook(uipath)
            print json.dumps(ExcelFile.sheet_names(),ensure_ascii=False,encoding='utf-8')#获取目标EXCEL文件sheet名
            sheet=ExcelFile.sheet_by_name(u'爱佑童心救助患儿明细')
            print sheet.name,sheet.nrows,sheet.ncols#打印sheet的名称，行数，列数
            cols=sheet.col_values(0)#获取整行或者整列的值#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            for row in range(len(cols)):
                reg_date = re.findall('\d{4}-\d{2}',cols[row].strip(),re.I|re.M)
                if reg_date:
                    print cols[row].strip()
                    a = row+1 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,0)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        item = {}
                        rows = sheet.row_values(row_content)  # 第三行内容
                        item['汇款时间'] = cols[row].strip().encode('utf-8')
                        item['救助费用'] = rows[1]
                        item['患儿编号'] = rows[2].encode('utf-8')
                        item['患儿姓名'] = rows[3].encode('utf-8')
                        item['性别'] = rows[4].encode('utf-8')
                        item['出生日期'] = rows[5].encode('utf-8')
                        item['汇款批次'] = rows[6].encode('utf-8')
                        item['救助医院'] = rows[7].encode('utf-8')
                        item['病种'] = rows[8].encode('utf-8')
                        item['治愈状态'] = rows[9].encode('utf-8')
                        item['手术费用'] = rows[10]
                        item['手术时间'] = rows[11].encode('utf-8')
                        item['入院时间'] = rows[12].encode('utf-8')
                        item['出院时间'] = rows[13].encode('utf-8')
                        item['所在省'] = rows[14].encode('utf-8')
                        item['患儿详细地址'] = rows[15].encode('utf-8')
                        item['联系电话'] =  (str(rows[16])).encode('utf-8')#rows[16].encode('utf-8') if type(rows[16]) == 'str' else
                        item['微信号'] = ""#rows[17].encode('utf-8')
                        item['捐赠人名字'] = path_key
                        item['项目类型'] = '爱佑童心救助患儿明细'
                        insert_sql = """insert into t_f_help_assist_detail_oa_copy1 (help_date,help_money,child_num,child_name,sex,child_birthday,batch,hosp_name,disease_name,remits_state,operation_money,operated_date,admission_datetime,discharge_datetime,province,detail_addr,contact_tel,extend_tel,donor_name,project_type) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"""%(item['汇款时间'],item['救助费用'],item['患儿编号'],item['患儿姓名'],item['性别'],item['出生日期'],item['汇款批次'],item['救助医院'],item['病种'],item['治愈状态'],item['手术费用'],item['手术时间'],item['入院时间'],item['出院时间'],item['所在省'],item['患儿详细地址'],item['联系电话'],item['微信号'],item['捐赠人名字'],item['项目类型'])
                        print insert_sql
                        cursor.execute(insert_sql)
                        conn.commit()


            sheet2=ExcelFile.sheet_by_name(u'童心分期患儿明细')
            print sheet2.name,sheet2.nrows,sheet2.ncols#打印sheet的名称，行数，列数
            cols=sheet2.col_values(0)#获取整行或者整列的值#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            for row in range(len(cols)):
                reg_date = re.findall('\d{4}-\d{2}',cols[row].strip(),re.I|re.M)
                if reg_date:
                    print cols[row].strip()
                    a = row+1 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet2,a,0)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        item = {}
                        rows = sheet2.row_values(row_content)  # 第三行内容
                        item['汇款时间'] = cols[row].strip().encode('utf-8')
                        item['救助费用'] = rows[1]
                        item['患儿编号'] = rows[2].encode('utf-8')
                        item['患儿姓名'] = rows[3].encode('utf-8')
                        item['性别'] = rows[4].encode('utf-8')
                        item['出生日期'] = rows[5].encode('utf-8')
                        item['汇款批次'] = rows[6].encode('utf-8')
                        item['救助医院'] = rows[7].encode('utf-8')
                        item['病种'] = rows[8].encode('utf-8')
                        item['手术费用'] = rows[9]
                        item['手术时间'] = rows[10].encode('utf-8')
                        item['入院时间'] = rows[11].encode('utf-8')
                        item['出院时间'] = rows[12].encode('utf-8')
                        item['所在省'] = rows[13].encode('utf-8')
                        item['患儿详细地址'] = rows[14].encode('utf-8')
                        item['联系电话'] = rows[15].encode('utf-8')
                        item['捐赠人名字'] = path_key
                        item['项目类型'] = '童心分期患儿明细'
                        insert_sql = """insert into t_f_help_assist_detail_oa_copy1 (help_date,help_money,child_num,child_name,sex,child_birthday,batch,hosp_name,disease_name,operation_money,operated_date,admission_datetime,discharge_datetime,province,detail_addr,contact_tel,donor_name,project_type) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"""%(item['汇款时间'],item['救助费用'],item['患儿编号'],item['患儿姓名'],item['性别'],item['出生日期'],item['汇款批次'],item['救助医院'],item['病种'],item['手术费用'],item['手术时间'],item['入院时间'],item['出院时间'],item['所在省'],item['患儿详细地址'],item['联系电话'],item['捐赠人名字'],item['项目类型'])
                        print insert_sql
                        cursor.execute(insert_sql)
                        conn.commit()

            sheet3=ExcelFile.sheet_by_name(u'童心门诊患儿明细')
            print sheet3.name,sheet3.nrows,sheet3.ncols#打印sheet的名称，行数，列数
            cols=sheet3.col_values(0)#获取整行或者整列的值#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            tongxin_date = '1111-11'
            for row in range(len(cols)):
                reg_date = re.findall('\d{4}-\d{2}',cols[row].strip(),re.I|re.M)
                if reg_date:
                    tongxin_date = get_merged_cells_value(sheet3,row,0)[0].encode('utf-8')
                else:
                    if tongxin_date == '1111-11':
                        pass
                    else:
                        item = {}
                        item['汇款时间'] = tongxin_date
                        item['救助费用'] = get_merged_cells_value(sheet3,row,1)[0]
                        flag = get_merged_cells_value(sheet3,row,0)[0]
                        if flag == u"汇总：":
                            pass
                        else:
                            item['患儿编号'] = get_merged_cells_value(sheet3,row,2)[0].encode('utf-8')
                            item['患儿姓名'] = get_merged_cells_value(sheet3,row,3)[0].encode('utf-8')
                            item['性别'] = get_merged_cells_value(sheet3,row,4)[0].encode('utf-8')
                            item['出生日期'] = get_merged_cells_value(sheet3,row,5)[0].encode('utf-8')
                            item['汇款批次'] = get_merged_cells_value(sheet3,row,6)[0].encode('utf-8')
                            item['救助医院'] = get_merged_cells_value(sheet3,row,7)[0].encode('utf-8')
                            item['病种'] = get_merged_cells_value(sheet3,row,8)[0].encode('utf-8')
                            item['门诊费用'] = get_merged_cells_value(sheet3,row,9)[0]
                            item['门诊时间'] = datetime(*xldate_as_tuple(get_merged_cells_value(sheet3, row, 10)[0], 0))
                            item['所在省'] = get_merged_cells_value(sheet3,row,11)[0].encode('utf-8')
                            item['患儿详细地址'] = get_merged_cells_value(sheet3,row,12)[0].encode('utf-8')
                            item['联系电话'] = get_merged_cells_value(sheet3,row,13)[0].encode('utf-8')
                            item['捐赠人名字'] = path_key
                            item['项目类型'] = '童心门诊患儿明细'
                            insert_sql = """insert into t_f_help_assist_detail_oa_copy1 (help_date,help_money,child_num,child_name,sex,child_birthday,batch,hosp_name,disease_name,mz_amt,mz_date,province,detail_addr,contact_tel,donor_name,project_type) values ('%s',%s,'%s','%s','%s','%s','%s','%s','%s',%s,'%s','%s','%s','%s','%s','%s')"""%(item['汇款时间'],item['救助费用'],item['患儿编号'],item['患儿姓名'],item['性别'],item['出生日期'],item['汇款批次'],item['救助医院'],item['病种'],item['门诊费用'],item['门诊时间'],item['所在省'],item['患儿详细地址'],item['联系电话'],item['捐赠人名字'],item['项目类型'])
                            print insert_sql
                            cursor.execute(insert_sql)
                            conn.commit()

            sheet4=ExcelFile.sheet_by_name(u'爱佑天使救助患儿明细')
            print sheet4.name,sheet4.nrows,sheet4.ncols#打印sheet的名称，行数，列数
            cols=sheet4.col_values(0)#获取整行或者整列的值#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            tianshi_date = '1111-11'
            for row in range(2,len(cols)-1):
                tianshi_reg_date = re.findall('\d{4}-\d{2}',cols[row].strip(),re.I|re.M)
                if tianshi_reg_date:
                    tianshi_date = get_merged_cells_value(sheet4,row,0)[0].encode('utf-8')
                else:
                    if tianshi_date == '1111-11':
                        pass
                    else:
                        item = {}
                        item['汇款月份'] = tianshi_date
                        item['患儿编号'] = get_merged_cells_value(sheet4,row,1)[0].encode('utf-8')
                        item['患儿姓名'] =get_merged_cells_value(sheet4,row,2)[0].encode('utf-8')
                        item['性别'] = get_merged_cells_value(sheet4,row,3)[0].encode('utf-8')
                        item['出生日期'] = get_merged_cells_value(sheet4,row,4)[0].encode('utf-8')
                        item['病种'] = get_merged_cells_value(sheet4,row,5)[0].encode('utf-8')
                        item['治疗期次'] = str(get_merged_cells_value(sheet4,row,6)[0]) if get_merged_cells_value(sheet4,row,6)[0] else ''.encode('utf-8')
                        item['救助费用'] = get_merged_cells_value(sheet4,row,7)[0]
                        item['入院时间'] = get_merged_cells_value(sheet4,row,8)[0].encode('utf-8')
                        item['出院时间'] = get_merged_cells_value(sheet4,row,9)[0].encode('utf-8')
                        item['救助医院'] = get_merged_cells_value(sheet4,row,10)[0].encode('utf-8')
                        item['汇款批次'] = get_merged_cells_value(sheet4,row,11)[0].encode('utf-8')
                        item['住院号'] = get_merged_cells_value(sheet4,row,12)[0].encode('utf-8')
                        item['患儿所在省'] = get_merged_cells_value(sheet4,row,13)[0].encode('utf-8')
                        item['患儿详细地址'] = get_merged_cells_value(sheet4,row,14)[0].encode('utf-8')
                        item['联系电话'] = get_merged_cells_value(sheet4,row,15)[0].encode('utf-8')
                        item['捐赠人名字'] = path_key
                        item['项目类型'] = '爱佑天使救助患儿明细'
                        insert_sql = """insert into t_f_help_assist_detail_oa_copy1 (help_date,child_num,child_name,sex,child_birthday,disease_name,period,help_money,admission_datetime,discharge_datetime,hosp_name,batch,AD_num,province,detail_addr,contact_tel,donor_name,project_type) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"""%(item['汇款月份'],item['患儿编号'],item['患儿姓名'],item['性别'],item['出生日期'],item['病种'],item['治疗期次'],item['救助费用'],item['入院时间'],item['出院时间'],item['救助医院'],item['汇款批次'],item['住院号'],item['患儿所在省'],item['患儿详细地址'],item['联系电话'],item['捐赠人名字'],item['项目类型'])
                        print insert_sql
                        cursor.execute(insert_sql)
                        conn.commit()

            sheet5=ExcelFile.sheet_by_name(u'爱佑天使特困患儿明细')
            print sheet5.name,sheet5.nrows,sheet5.ncols#打印sheet的名称，行数，列数
            cols=sheet5.col_values(0)#获取整行或者整列的值#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            tianshi_date = '1111-11'
            for row in range(2,len(cols)-1):
                tianshi_reg_date = re.findall('\d{4}-\d{2}',cols[row].strip(),re.I|re.M)
                if tianshi_reg_date:
                    tianshi_date = get_merged_cells_value(sheet5,row,0)[0].encode('utf-8')
                else:
                    if tianshi_date == '1111-11':
                        pass
                    else:
                        item = {}
                        item['汇款月份'] = tianshi_date
                        item['患儿编号'] = get_merged_cells_value(sheet5,row,1)[0].encode('utf-8')
                        item['患儿姓名'] = get_merged_cells_value(sheet5,row,2)[0].encode('utf-8')
                        item['性别'] = get_merged_cells_value(sheet5,row,3)[0].encode('utf-8')
                        item['出生日期'] = get_merged_cells_value(sheet5,row,4)[0].encode('utf-8')
                        item['病种'] = get_merged_cells_value(sheet5,row,5)[0].encode('utf-8')
                        item['治疗期次'] = get_merged_cells_value(sheet5,row,6)[0]
                        item['救助费用'] = get_merged_cells_value(sheet5,row,7)[0]
                        item['入院时间'] = get_merged_cells_value(sheet5,row,8)[0].encode('utf-8')
                        item['出院时间'] = get_merged_cells_value(sheet5,row,9)[0].encode('utf-8')
                        item['救助医院'] = get_merged_cells_value(sheet5,row,10)[0].encode('utf-8')
                        item['住院号'] = get_merged_cells_value(sheet5,row,11)[0].encode('utf-8')
                        item['患儿所在省'] = get_merged_cells_value(sheet5,row,12)[0].encode('utf-8')
                        item['患儿详细地址'] = get_merged_cells_value(sheet5,row,13)[0].encode('utf-8')
                        item['联系电话'] = get_merged_cells_value(sheet5,row,14)[0].encode('utf-8')
                        item['捐赠人名字'] = path_key
                        item['项目类型'] = '爱佑天使特困患儿明细'
                        insert_sql = """insert into t_f_help_assist_detail_oa_copy1 (help_date,child_num,child_name,sex,child_birthday,disease_name,period,help_money,admission_datetime,discharge_datetime,hosp_name,AD_num,province,detail_addr,contact_tel,donor_name,project_type) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"""%(item['汇款月份'],item['患儿编号'],item['患儿姓名'],item['性别'],item['出生日期'],item['病种'],item['治疗期次'],item['救助费用'],item['入院时间'],item['出院时间'],item['救助医院'],item['住院号'],item['患儿所在省'],item['患儿详细地址'],item['联系电话'],item['捐赠人名字'],item['项目类型'])
                        print insert_sql
                        cursor.execute(insert_sql)
                        conn.commit()

            sheet6=ExcelFile.sheet_by_name(u'爱佑天使救助人数')
            print sheet6.name,sheet6.nrows,sheet6.ncols#打印sheet的名称，行数，列数
            cols=sheet6.col_values(0)#获取整行或者整列的值#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            for row in range(2,len(cols)-1):
                item = {}
                item['患儿编号'] = get_merged_cells_value(sheet6,row,0)[0].encode('utf-8')
                item['患儿姓名'] = get_merged_cells_value(sheet6,row,1)[0].encode('utf-8')
                item['性别'] = get_merged_cells_value(sheet6,row,2)[0].encode('utf-8')
                item['出生日期'] = get_merged_cells_value(sheet6,row,3)[0].encode('utf-8')
                item['病种'] = get_merged_cells_value(sheet6,row,4)[0].encode('utf-8')
                item['救助医院'] = get_merged_cells_value(sheet6,row,5)[0].encode('utf-8')
                item['住院号'] = get_merged_cells_value(sheet6,row,6)[0].encode('utf-8')
                item['患儿所在省'] = get_merged_cells_value(sheet6,row,7)[0].encode('utf-8')
                item['患儿详细地址'] = get_merged_cells_value(sheet6,row,8)[0].encode('utf-8')
                item['联系电话'] = get_merged_cells_value(sheet6,row,9)[0].encode('utf-8')
                item['治疗期次'] = str(sheet6.cell_value(row,10)) if sheet6.cell_value(row,10) else ''.encode('utf-8')
                item['入院时间'] = sheet6.cell_value(row,11).encode('utf-8')
                item['出院时间'] = sheet6.cell_value(row,12).encode('utf-8')
                item['救助费用'] = sheet6.cell_value(row,13)
                item['捐赠人名字'] = path_key
                item['项目类型'] = '爱佑天使救助人数'
                if item['患儿姓名']:
                    insert_sql = """insert into t_f_help_assist_detail_oa_copy1 (child_num,child_name,sex,child_birthday,disease_name,hosp_name,AD_num,province,detail_addr,contact_tel,period,admission_datetime,discharge_datetime,help_money,donor_name,project_type) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"""%(item['患儿编号'],item['患儿姓名'],item['性别'],item['出生日期'],item['病种'],item['救助医院'],item['住院号'],item['患儿所在省'],item['患儿详细地址'],item['联系电话'],item['治疗期次'],item['入院时间'],item['出院时间'],item['救助费用'],item['捐赠人名字'],item['项目类型'])
                    print insert_sql
                    cursor.execute(insert_sql)
                    conn.commit()

            sheet7=ExcelFile.sheet_by_name(u'爱佑晨星救助明细')
            print sheet7.name,sheet7.nrows,sheet7.ncols#打印sheet的名称，行数，列数
            cols=sheet7.col_values(0)#获取整行或者整列的值#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            cengxing_date = '1111-11'
            for row in range(2,len(cols)-1):
                cengxing_reg_date = re.findall('\d{4}-\d{2}',cols[row].strip(),re.I|re.M)
                if cengxing_reg_date:
                    cengxing_date = get_merged_cells_value(sheet7,row,0)[0].encode('utf-8')
                else:
                    if cengxing_date == '1111-11':
                        pass
                    else:
                        item = {}
                        item['汇款月份'] = cengxing_date
                        item['患儿编号'] = get_merged_cells_value(sheet7,row,1)[0].encode('utf-8')
                        item['患儿姓名'] = get_merged_cells_value(sheet7,row,2)[0].encode('utf-8')
                        item['性别'] = get_merged_cells_value(sheet7,row,3)[0].encode('utf-8')
                        item['出生日期'] = get_merged_cells_value(sheet7,row,4)[0].encode('utf-8')
                        item['治疗期次'] = str(get_merged_cells_value(sheet7,row,5)[0]) if get_merged_cells_value(sheet7,row,5)[0] else ''.encode('utf-8')
                        item['救助费用'] = get_merged_cells_value(sheet7,row,6)[0]
                        item['病种'] = get_merged_cells_value(sheet7,row,7)[0].encode('utf-8')
                        item['病种名称'] = get_merged_cells_value(sheet7,row,8)[0].encode('utf-8')
                        item['救助医院'] = get_merged_cells_value(sheet7,row,9)[0].encode('utf-8')
                        item['住院号'] = get_merged_cells_value(sheet7,row,10)[0].encode('utf-8')
                        item['患儿所在省'] = get_merged_cells_value(sheet7,row,11)[0].encode('utf-8')
                        item['患儿详细地址'] = get_merged_cells_value(sheet7,row,12)[0].encode('utf-8')
                        item['联系电话'] = get_merged_cells_value(sheet7,row,13)[0].encode('utf-8')
                        item['入院时间'] = get_merged_cells_value(sheet7,row,14)[0].encode('utf-8')
                        item['出院时间'] = get_merged_cells_value(sheet7,row,15)[0].encode('utf-8')
                        item['捐赠人名字'] = path_key
                        item['项目类型'] = '爱佑晨星救助明细'
                        insert_sql = """insert into t_f_help_assist_detail_oa_copy1 (remits_date,child_num,child_name,sex,child_birthday,bz,disease_name,province,detail_addr,contact_tel,hosp_name,period,admission_datetime,discharge_datetime,help_money,donor_name,project_type) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"""%(item['汇款月份'],item['患儿编号'],item['患儿姓名'],item['性别'],item['出生日期'],item['病种'],item['病种名称'],item['患儿所在省'],item['患儿详细地址'],item['联系电话'],item['救助医院'],item['治疗期次'],item['入院时间'],item['出院时间'],item['救助费用'],item['捐赠人名字'],item['项目类型'])
                        print insert_sql
                        cursor.execute(insert_sql)
                        conn.commit()


            sheet8 = ExcelFile.sheet_by_name(u'爱佑晨星救助人数')
            print sheet8.name, sheet8.nrows, sheet8.ncols  # 打印sheet的名称，行数，列数
            cols = sheet8.col_values(0)  # 获取整行或者整列的值#第二列内容
            print json.dumps(cols, encoding='utf-8', ensure_ascii=False)
            for row in range(2, len(cols)):
                item = {}
                item["患儿编号"] = get_merged_cells_value(sheet8, row, 0)[0].encode('utf-8')
                item["患儿姓名"] = get_merged_cells_value(sheet8, row, 1)[0].encode('utf-8')
                item["性别"] = get_merged_cells_value(sheet8, row, 2)[0].encode('utf-8')
                item["出生日期"] = get_merged_cells_value(sheet8, row, 3)[0].encode('utf-8')
                item["病种"] = get_merged_cells_value(sheet8, row, 4)[0].encode('utf-8')
                item["病种名称"] = get_merged_cells_value(sheet8, row, 5)[0].encode('utf-8')
                item["患儿所在省"] = get_merged_cells_value(sheet8, row, 6)[0].encode('utf-8')
                item["患儿详细地址"] = get_merged_cells_value(sheet8, row, 7)[0].encode('utf-8')
                item["联系电话"] = get_merged_cells_value(sheet8, row, 8)[0].encode('utf-8')
                item["救助医院"] = get_merged_cells_value(sheet8, row, 9)[0].encode('utf-8')
                item["治疗期次"] = str(get_merged_cells_value(sheet8, row, 10)[0])
                item["入院时间"] = get_merged_cells_value(sheet8, row, 11)[0].encode('utf-8')
                item["出院时间"] = get_merged_cells_value(sheet8, row, 12)[0].encode('utf-8')
                item["救助费用"] =  get_merged_cells_value(sheet8, row, 13)[0]
                item["捐赠人名字"] = path_key
                item['项目类型'] = '爱佑晨星救助人数'
                print json.dumps(item,ensure_ascii=False,encoding='utf-8')
                insert_sql = """insert into t_f_help_assist_detail_oa_copy1 (child_num,child_name,sex,child_birthday,bz,disease_name,province,detail_addr,contact_tel,hosp_name,period,admission_datetime,discharge_datetime,help_money,donor_name,project_type) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s',%s,'%s','%s',%s,'%s','%s')""" % (item["患儿编号"],item["患儿姓名"],item["性别"],item["出生日期"], item["病种"],item["病种名称"],item["患儿所在省"], item["患儿详细地址"], item["联系电话"], item["救助医院"], item["治疗期次"], item["入院时间"], item["出院时间"],item["救助费用"],item['捐赠人名字'], item['项目类型'])
                print insert_sql
                cursor.execute(insert_sql)
                conn.commit()

if __name__ =='__main__':
    read_excel(unicode(r'C:\Users\pig\Desktop\donor', 'utf-8'))
    cursor.close()
    conn.close()
    print "finish"