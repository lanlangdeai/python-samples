# -*- coding: utf-8 -*-
# @Time  : 2019/10/16
# @Author: X-Wolf
# @Describe  通过公司名称获取营业状态
import json
import re
import urllib

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

    res = requests.get(url, headers=headers, verify=None)
    res.encoding = 'utf-8'

    soup = BeautifulSoup(res.text,'lxml')
    result = soup.find('tbody',id="search-result")
    id = result.find('input', attrs={'name':'batchPostcard'})['value']
    print(id)
    # print(result)
    status = result.find('span',class_="nstatus").text
    print(status)
    # print(res.text)

# 获取公司注销或吊销时间
def fetchCompanyRevokeDate(url):
    res = requests.get(url, headers=headers, verify=None)
    res.encoding = 'utf-8'
    soup = BeautifulSoup(res.text, 'lxml')
    ret = soup.find('span', attrs={'class': 'ntag text-danger tooltip-br'})
    print(ret.prettify())
    title = ret.attrs['title']
    pattern = re.compile(r'<[^>]+>', re.S)
    result = pattern.sub('', title)
    print(result)
    pattern2 = re.compile(r'(\d{4}-\d{1,2}-\d{1,2})', re.S | re.M)
    date = pattern2.search(result).group(0)
    print(date)
    text = ret.get_text()
    print(text)


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
    # fetchCompanyRevokeDate(url)
    name = '北京优你课教育科技有限公司'
    fetchCompanyStatusAndIdByName(name)

if __name__ == '__main__':
    run()