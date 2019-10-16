# -*- coding: utf-8 -*-
# @Time  : 2019/10/16
# @Author: X-Wolf
# @Describe  通过公司名称获取营业状态
import json
import re
import urllib
import xlwt
import xlrd
import time

import requests
from bs4 import BeautifulSoup

# 先找到运营状态
# 吊销或注销的查询具体时间

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36 SocketLog(tabid=177&client_id=)'
}


# 通过公司名称获取营业状态与唯一ID
def fetchCompanyStatusAndIdByName(name):

    # 首页需要获取到cookie,再请求
    key = urllib.parse.quote(name)
    url = 'https://www.qichacha.com/search?key='+key
    headers['Referer'] = url
    headers['Cookie'] = json.dumps(getCookie())
    print(url)
    res = requests.get(url, headers=headers, verify=None)
    res.encoding = 'utf-8'
    print('获取公司唯一标识 状态：{} 数据：{}'.format(res.status_code,res.text))

    soup = BeautifulSoup(res.text,'lxml')
    result = soup.find('tbody',id="search-result")
    id = result.find('input', attrs={'name':'batchPostcard'})['value']
    # print(id)
    # print(result)
    status = result.find('span',class_="nstatus").text
    # print(status)
    return id

# 获取公司注销或吊销时间
def fetchCompanyRevokeDate(url):
    res = requests.get(url, headers=headers, verify=None)
    res.encoding = 'utf-8'
    print('获取时间 状态：{} 返回数据:{}'.format(res.status_code, res.text))
    soup = BeautifulSoup(res.text, 'lxml')
    ret = soup.find('span', attrs={'class': 'ntag text-danger tooltip-br'})
    # print(ret.prettify())
    title = ret.attrs['title']
    pattern = re.compile(r'<[^>]+>', re.S)
    result = pattern.sub('', title)
    print(result)

    pattern2 = re.compile(r'(\d{4}-\d{1,2}-\d{1,2})', re.S | re.M)
    try:
        date = pattern2.search(result).group(0)
    except:
        date = '无'
    # print(date)
    # text = ret.get_text()
    # print(text)
    return date

def getCookie():
    host_url = 'https://www.qichacha.com'
    res = requests.get(host_url, headers=headers)
    res.encoding = 'utf-8'
    # print(res.text,res.status_code)
    cookie = dict()
    for key, value in res.cookies.items():
        cookie[key] = value
    return cookie

def run():
    # url = 'https://www.qichacha.com/firm_f31f957151a40454a92f1b8370d9083e' # 吊销
    # url = 'https://www.qichacha.com/firm_990050ea0fda94fbd508a80647aa1589' # 注销

    #读取文件中数据，写入到xls文件中
    dest_file = '公司信息.xls'
    wt = xlwt.Workbook()
    sheet1 = wt.add_sheet('Sheet1', cell_overwrite_ok=True)

    src_file = '公司信息10-1.xls'
    work = xlrd.open_workbook(src_file)
    sheet = work.sheet_by_name('Sheet1')

    rows = sheet.nrows
    rows = 12

    cols = sheet.ncols
    for i in range(0,rows):
        print(f'当前行：{i}')
        item = sheet.row_values(i)
        if i == 0:
            item.append('注销或吊销时间')
        else:
            name = item[0]
            status = item[1]
            print(f'公司名称：{name}')
            print(f'经营状态：{status}')
            if status == '注销' or status == '吊销':
                flag = fetchCompanyStatusAndIdByName(name)
                print('唯一标示：'+flag)
                url = 'https://www.qichacha.com/firm_'+flag
                date = fetchCompanyRevokeDate(url)
                print('注销或吊销日期：'+date)
                item.append(date)
            else:
                item.append('')

        for j in range(0, cols+1):
            print(f'当前字段：{j}')
            sheet1.write(i,j,item[j])

        time.sleep(1)

    wt.save(dest_file)

    print('finish')

if __name__ == '__main__':
    run()