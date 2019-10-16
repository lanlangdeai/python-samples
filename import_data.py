# -*- coding: utf-8 -*-
# @Time  : 2019/10/16
# @Author: X-Wolf
# @Describe 导数据操作
import pymysql
import xlrd

db = pymysql.Connect(
    host='localhost',
    port=3306,
    user='root',
    passwd='',
    db = 'test',
    charset='utf8',
    # curorclass=pymysql.cursors.DictCursor
)
cursor = db.cursor()


#将数据从xls导入MySQL
def xlsToMysql(file_name):
    work = xlrd.open_workbook(file_name)
    sheet = work.sheet_by_name('Sheet1')
    for i in range(1,sheet.nrows):
        row = sheet.row_values(i)
        # print(row)
        insert(row)

# 将数据插入数据库
def insert(data):
    # print(data[0],data[20])
    # exit(0)
    fields = 'company_name, manage_status, legal_person, registered_assets, ' \
             'establish_date, province, phone, other_phones, email, credit_code, ' \
             'identity_code, registered_code, agency_number, insured_number'
    # fields = 'company_name, manage_status'
    sql = "INSERT INTO `qichacha_company` (%s) VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s') " % (
    # sql = "INSERT INTO `qichacha_company` (%s) VALUES('%s','%s') " % (
        fields,
        data[0],
        data[1],
        data[2],
        data[3],
        data[4],
        data[5],
        data[6],
        data[7],
        data[8],
        data[9],
        data[10],
        data[11],
        data[12],
        str(data[13]),
        # data[14],
        # data[15],
        # data[16],
        # data[17],
        # data[18],
        # data[19],
        # int(data[20]),
    )
    # print(sql)
    # exit(0)
    try:
        cursor.execute(sql)
        ret = db.commit()
        return True
    except Exception as e:
        db.rollback()
        print('发生错误',e)
        return False


def run():
    for i in range(1,2):
        file_name = f'公司信息{i}-1.xls'
        xlsToMysql(file_name)


if __name__ == '__main__':
    run()