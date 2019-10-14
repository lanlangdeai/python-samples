#!/usr/bin/env python
# @Time  : 2019/10/13 17:09
# @Author: X-Wolf
# @Describe 将西瓜数据从mongo到处到xlsx文件中
import pymongo
import xlwt
import xlrd

mg = pymongo.MongoClient(host='localhost', port=27017)
db = mg.xiguadata
collection = db.zhuti

'''
    将数据导出到xlsx文件
'''
def exportToXlsx():

    print('task begin')
    # 将数据/5000进行分割到多个xlsx文件中
    page, limit = 1, 5000
    while True:
        print('当前页面：', page)
        # 打开文件
        # xls = xlwt.Workbook()
        # sheet1 = xls.add_sheet('Sheet1', cell_overwrite_ok=True)
        # sheet1.write(0, 0, '批量查询')
        offset = (page-1)*limit
        result = collection.find({}, {'compname': 1, 'account_num': 1})\
                .sort('account_num', pymongo.DESCENDING)\
                .skip(offset).limit(limit)
        if result and result.count() > offset:
            i = 1
            for item in result:
                print('读取数据i:', i)
                # sheet1.write(i, 0, item['compname'])
                i += 1
        else:
            print('Done')
            break

        # xls.save(str(page)+'.xlsx')


        page += 1

def test():
    # 山西老乡科技有限公司
    result = collection.find_one({'compname': {'$regex': '老乡'}}, {'compname': 1, 'account_num': 1})
    # total = result.count()
    print(result)

def addAccountNumToData(file):
    # 读取文件内容，并添加公众号数据字段
    workbook = xlrd.open_workbook(file)
    sheet_name = workbook.sheet_names()[0]
    print('sheet 名称',sheet_name)
    sheet = workbook.sheet_by_name(sheet_name)
    print('行数：', sheet.nrows)
    print('列数：', sheet.ncols)

    xls = xlwt.Workbook()
    sheet1 = xls.add_sheet('Sheet1', cell_overwrite_ok=True)
    for i in range(1, sheet.nrows):
        row = sheet.row_values(i)
        print('当前行数：', i)
        if i == 1:
            row.append('公众号数量')
        else:
            # 查询数量并追加
            compname = row[0]
            print('公司名称', compname)
            item = collection.find_one({'compname': compname}, {'account_num': 1})
            # print('公司信息：', item['account_num'])
            row.append(item['account_num'] if item != None else 0)
            print(row)
            # break
        #写入新XLS
        for j in range(0, len(row)):
            # print('处理列。。。')
            sheet1.write(i-1, j, row[j]) # 行，列，值
        # break


    #写入文件中
    xls.save('./principal-data/公司信息'+file)

if __name__ == '__main__':
    # exportToXlsx()
    # test()
    # exit(0)
    xls_files = ['2-1.xls','3-1.xls','4-1.xls','5-1.xls','6-1.xls','7-1.xls','8-1.xls','9-1.xls','10-1.xls']
    for file in xls_files:
        addAccountNumToData(file)