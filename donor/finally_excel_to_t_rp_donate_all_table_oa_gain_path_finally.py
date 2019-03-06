# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
import xlrd
import pymysql
from datetime import date,datetime
import json
import os
"""
汇总
"""
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
            # print(file_name)
            (filepath, tempfilename) = os.path.split(file_name)
            # print filepath, tempfilename
            (filename, extension) = os.path.splitext(tempfilename)
            # print filename, extension
            print filename,file_name
            path_key = filename
            path_value = file_name
        # for path_key,path_value in path_dict.items():
            # inpath = r'D:\desktop\14_数据仓库组\捐赠人明细录入文件解析程序到数据库\顺丰（截止至11-22）.xlsx'  # 文件位置
            inpath = path_value
            # uipath = unicode(inpath, "utf8")
            uipath = path_value
            ExcelFile=xlrd.open_workbook(uipath)
            #获取目标EXCEL文件sheet名
            print json.dumps(ExcelFile.sheet_names(),ensure_ascii=False,encoding='utf-8')
            sheet=ExcelFile.sheet_by_name(u'捐赠及支出额度汇总')
            #打印sheet的名称，行数，列数
            print sheet.name,sheet.nrows,sheet.ncols
            #获取整行或者整列的值
            cols=sheet.col_values(1)#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            for row in range(len(cols)):
                if cols[row].strip() == u'累计捐赠额度汇总及明细':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        if rows[3]:
                            rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                            # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                            insert_sql = "insert into t_f_donate_account_detail_oa_copy1 (donor,donate_dt,donate_amt,donate_subject_name) values ('%s','%s','%s','%s')"%(rows[1],rows[2],rows[3],rows[4])
                            print insert_sql
                            cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑童心手术累计支出额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_detail_oa_copy1 (donor,donate_date,help_amt,help_child_cnt,project_cd,project_type) values ('%s','%s',%s,%s,'%s','%s')"%(rows[1],rows[2],rows[3],rows[4],'ASSIST_AYTX','3')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑童心手术待付款额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_nopay_detail_oa_copy1 (donor,help_child_cnt,help_amt,project_cd,project_type) values ('%s','%s','%s','%s','%s')"%(rows[1],rows[2],rows[3],'ASSIST_AYTX','3')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑童心分期累计支出额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_detail_oa_copy1 (donor,donate_date,help_amt,help_child_cnt,project_cd,project_type) values ('%s','%s',%s,%s,'%s','%s')"%(rows[1],rows[2],rows[3],rows[4],'ASSIST_AYTX','1')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑童心分期待付款额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_nopay_detail_oa_copy1 (donor,help_amt,help_child_cnt,project_cd,project_type) values ('%s',%s,%s,'%s','%s')"%(rows[1],rows[2],rows[3],'ASSIST_AYTX','1')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑童心门诊累计支出额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_detail_oa_copy1 (donor,donate_date,help_amt,help_child_cnt,project_cd,project_type) values ('%s','%s',%s,%s,'%s','%s')"%(rows[1],rows[2],rows[3],rows[4],'ASSIST_AYTX','2')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑童心门诊待付款额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_nopay_detail_oa_copy1 (donor,help_child_cnt,help_amt,project_cd,project_type) values ('%s',%s,%s,'%s','%s')"%(rows[1],rows[2],rows[3],'ASSIST_AYTX','2')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑天使累计支出额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_detail_oa_copy1 (donor,donate_date,help_amt,help_child_cnt,project_cd) values ('%s','%s',%s,%s,'%s')"%(rows[1],rows[2],rows[3],rows[4],'ASSIST_AYTS')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑天使特困患儿累计支出额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_detail_oa_copy1 (donor,donate_date,help_amt,help_child_cnt,project_cd,project_type) values ('%s','%s',%s,%s,'%s','%s')"%(rows[1],rows[2],rows[3],rows[4],'ASSIST_AYTS','5')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑天使待付款额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_nopay_detail_oa_copy1 (donor,help_child_cnt,help_amt,project_cd) values ('%s',%s,%s,'%s')"%(rows[1],rows[2],rows[3],'ASSIST_AYTS')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑天使特困患儿待付款额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_nopay_detail_oa_copy1 (donor,help_child_cnt,help_amt,project_cd,project_type) values ('%s',%s,%s,'%s','%s')"%(rows[1],rows[2],rows[3],'ASSIST_AYTS','5')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑天使资助未完成数量':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_process_month_detail_oa_copy1 (donor,child_cnt,amt,left_amt,project_cd) values ('%s',%s,%s,%s,'%s')"%(rows[1],rows[2],rows[3],rows[4],'ASSIST_AYTS')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑晨星累计支出额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_detail_oa_copy1 (donor,donate_date,help_amt,help_child_cnt,project_cd) values ('%s','%s',%s,%s,'%s')"%(rows[1],rows[2],rows[3],rows[4],'ASSIST_AYCX')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑晨星待付款额度汇总':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_month_nopay_detail_oa_copy1 (donor,help_amt,help_child_cnt,project_cd) values ('%s',%s,%s,'%s')"%(rows[1],rows[2],rows[3],'ASSIST_AYCX')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'爱佑晨星资助未完成数量':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_process_month_detail_oa_copy1 (donor,child_cnt,amt,left_amt,project_cd) values ('%s',%s,%s,%s,'%s')"%(rows[1],rows[2],rows[3],rows[4],'ASSIST_AYCX')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'童心单次手术退款':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_refund_month_detail_oa_copy1 (donor,child_cnt,amt,project_cd,project_type,refund_flag) values ('%s',%s,%s,'%s','%s','%s')" % (rows[1], rows[2], rows[3], 'ASSIST_AYTX', '3', 'R')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'童心门诊退款':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_refund_month_detail_oa_copy1 (donor,child_cnt,amt,project_cd,project_type,refund_flag) values ('%s',%s,%s,'%s','%s','%s')" % (rows[1], rows[2], rows[3], 'ASSIST_AYTX', '2', 'R')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'童心分期退款':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_refund_month_detail_oa_copy1 (donor,child_cnt,amt,project_cd,project_type,refund_flag) values ('%s',%s,%s,'%s','%s','%s')" % (rows[1], rows[2], rows[3], 'ASSIST_AYTX', '1', 'R')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'童心补款':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_refund_month_detail_oa_copy1 (donor,child_cnt,amt,project_cd,refund_flag) values ('%s',%s,%s,'%s','%s')" % (rows[1], rows[2], rows[3], 'ASSIST_AYTX', 'S')
                        print insert_sql
                        cursor.execute(insert_sql)
                if cols[row].strip() == u'天使退款':
                    print cols[row].strip()
                    a = row+2 #用来取行数头（a+2）
                    b = get_merged_cells_value(sheet,a,1)  #用来取行数尾（行数头，行数尾，列数头，列数尾）
                    c = b[1][1]  #获取合并单元格尾部坐标
                    for row_content in range(a,c):
                        rows = sheet.row_values(row_content)  # 第三行内容
                        rows[1] = get_merged_cells_value(sheet,row_content,1)[0] #获取合并单元格内容
                        # print json.dumps(rows,ensure_ascii=False,encoding='utf-8')  #获取整列内容
                        insert_sql = "insert into t_rp_donate_refund_month_detail_oa_copy1 (donor,child_cnt,amt,project_cd,project_type,refund_flag) values ('%s',%s,%s,'%s','%s','%s')" % (rows[1], rows[2], rows[3], 'ASSIST_AYTS','4', 'R')
                        print insert_sql
                        cursor.execute(insert_sql)
            sheet2=ExcelFile.sheet_by_name(u'退补款明细')
            #打印sheet的名称，行数，列数
            print sheet2.name,sheet2.nrows,sheet2.ncols
            #获取整行或者整列的值
            cols=sheet2.col_values(0)#第二列内容
            print json.dumps(cols,encoding='utf-8',ensure_ascii=False)
            for row in range(len(cols)):
                if cols[row].strip() == u'手术退款明细':
                    print cols[row].strip()
                    a = row + 2  # 用来取行数头（a+2）
                    flag = True
                    while flag:
                        rows = sheet2.row_values(a)  # 第三行内容
                        if rows[1]:
                            insert_sql = "insert into t_f_refund_detail_oa_copy1 (donor,child_num,child_name,help_money,refund_amt,refund_reason,refund_flag) values ('%s','%s','%s',%s,%s,'%s','%s')"%(path_key,rows[0],rows[1],rows[2],rows[3],rows[4],u'R')
                            print insert_sql
                            cursor.execute(insert_sql)
                            a += 1
                        else:
                            flag = False
                if cols[row].strip() == u'门诊退款明细':
                    print cols[row].strip()
                    a = row + 2  # 用来取行数头（a+2）
                    flag = True
                    while flag:
                        rows = sheet2.row_values(a)  # 第三行内容
                        if rows[1]:
                            insert_sql = "insert into t_f_refund_detail_oa_copy1 (donor,child_num,child_name,help_money,refund_amt,refund_reason,refund_flag) values ('%s','%s','%s',%s,%s,'%s','%s')" % (path_key, rows[0], rows[1], rows[2], rows[3], rows[4], u'R')
                            print insert_sql
                            cursor.execute(insert_sql)
                            a += 1
                        else:
                            flag = False
                if cols[row].strip() == u'分期退款明细':
                    print cols[row].strip()
                    a = row + 2  # 用来取行数头（a+2）
                    flag = True
                    while flag:
                        rows = sheet2.row_values(a)  # 第三行内容
                        if rows[1]:
                            insert_sql = "insert into t_f_refund_detail_oa_copy1 (donor,child_num,child_name,help_money,refund_amt,refund_reason,refund_flag) values ('%s','%s','%s',%s,%s,'%s','%s')" % (path_key, rows[0], rows[1], rows[2], rows[3], rows[4], u'R')
                            print insert_sql
                            cursor.execute(insert_sql)
                            a += 1
                        else:
                            flag = False
                if cols[row].strip() == u'童心手术补款明细':
                    print cols[row].strip()
                    a = row + 2  # 用来取行数头（a+2）
                    flag = True
                    while flag:
                        rows = sheet2.row_values(a)  # 第三行内容
                        if rows[1]:
                            insert_sql = "insert into t_f_refund_detail_oa_copy1 (donor,child_num,child_name,help_money,refund_amt,refund_reason,refund_flag) values ('%s','%s','%s',%s,%s,'%s','%s')" % (path_key, rows[0], rows[1], rows[2], rows[3], rows[4], u'S')
                            print insert_sql
                            cursor.execute(insert_sql)
                            a += 1
                        else:
                            flag = False
                if cols[row].strip() == u'天使退款明细':
                    print cols[row].strip()
                    a = row + 2  # 用来取行数头（a+2）
                    flag = True
                    while flag:
                        rows = sheet2.row_values(a)  # 第三行内容
                        if rows[1]:
                            insert_sql = "insert into t_f_refund_detail_oa_copy1 (donor,child_num,child_name,help_money,refund_amt,refund_reason,refund_flag) values ('%s','%s','%s',%s,%s,'%s','%s')" % (path_key, rows[0], rows[1], rows[2], rows[3], rows[4], u'R')
                            print insert_sql
                            cursor.execute(insert_sql)
                            a += 1
                        else:
                            flag = False
                if cols[row].strip() == u'晨星退款明细':
                    print cols[row].strip()
                    a = row + 2  # 用来取行数头（a+2）
                    flag = True
                    while flag:
                        if len(cols) > a:
                            rows = sheet2.row_values(a)  # 第三行内容
                            if rows[1]:
                                insert_sql = "insert into t_f_refund_detail_oa_copy1 (donor,child_num,child_name,help_money,refund_amt,refund_reason,refund_flag) values ('%s','%s','%s',%s,%s,'%s','%s')" % (path_key, rows[0], rows[1], rows[2], rows[3], rows[4], u'R')
                                print insert_sql
                                cursor.execute(insert_sql)
                                a += 1
                        else:
                            flag = False
        conn.commit()

if __name__ =='__main__':
    read_excel(unicode(r'C:\Users\pig\Desktop\donor', 'utf-8'))
    cursor.close()
    conn.close()
    print "finish"