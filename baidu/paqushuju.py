# -*- coding:utf-8 -*-
# paqushuju.py
from selenium import webdriver
import time
import xlrd
import xlsxwriter
import time
import os

# 读取 query2文件夹下所有的文件，list
query_path = r'D:\babytree\codetest\zuoye\pachong\query2'
wenjianlist = os.listdir(query_path)

# 给生成的新的文件，命名
namelist = []
PathList = []
for single in wenjianlist:
    newsingle = single.split('.')[0] + '的搜索结果.xlsx'
    namelist.append(newsingle)
    # 要读取的每个文件的位置
    PathList.append(r'D:\babytree\codetest\zuoye\pachong\query2' + "\\" + single)

for namei in range(len(wenjianlist)):
    time.sleep(5)

    ###############################################
    # 初始化浏览器
    browser = webdriver.Chrome(r'D:\python3\Lib\chromedriver.exe')
    browser.get('https://www.baidu.com/')
    browser.maximize_window()
    browser.find_element_by_xpath('//*[@id="kw"]').send_keys(u'怎么样看早孕试纸11')
    browser.find_element_by_xpath('//*[@id="su"]').click()
    time.sleep(1)
    browser.find_element_by_id('kw').clear()

    ##################################################
    # 创建一个workbook 设置编码
    workbook = xlsxwriter.Workbook(namelist[namei])

    #####################################################
    # 数据初始化

    # 文件地址+名称
    thePath = PathList[namei]

    # 读取第几列，这是一个列表
    dataColumn = []
    # 读取名称是第几列的数据
    dataCouName = '点击位置'

    # 打开文件
    data = xlrd.open_workbook(thePath)

    # 所有的sheet页，是一个list
    sheetlist = data.sheet_names()
    for sheetNamei in range(len(sheetlist)):
        # 创建一个worksheet
        worksheet = workbook.add_worksheet(sheetlist[sheetNamei])

        # 循环读取的sheet页

        # 通过文件名获得工作表,获取工作表
        table = data.sheet_by_name(sheetlist[sheetNamei])
        # 获取第一列数据
        # 第一列数据，是个list，然后干掉列表第一个
        fistCol = table.col_values(0)
        if fistCol[0] == 'QUERY' or fistCol[0] == 'yunyu':
            fistCol.pop(0)
        else:
            pass

        for searchKyi in range(len(fistCol)):
            browser.implicitly_wait(15)
            try:
                browser.find_element_by_id('kw').send_keys(fistCol[searchKyi])

            except:
                time.sleep(2)
                browser.find_element_by_id('kw').send_keys(fistCol[searchKyi])

            browser.implicitly_wait(15)
            try:
                browser.find_element_by_id('su').click()
            except:
                time.sleep(2)
                browser.find_element_by_id('su').click()
            browser.implicitly_wait(15)
            browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(6)

            try:
                a = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[1]/th[1]/a').text
            except:
                a = '无相关数据'
            try:
                b = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[1]/th[2]/a').text
            except:
                b = '无相关数据'
            try:
                c = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[1]/th[3]/a').text
            except:
                c = '无相关数据'
            try:
                d = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[2]/th[1]/a').text
            except:
                d = '无相关数据'
            try:

                e = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[2]/th[2]/a').text
            except:
                e = '无相关数据'

            try:
                f = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[2]/th[3]/a').text
            except:
                f = '无相关数据'
            try:
                g = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[3]/th[1]/a').text

            except:
                g = '无相关数据'
            try:
                h = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[3]/th[2]/a').text
            except:
                h = '无相关数据'
            try:
                i = browser.find_element_by_xpath('//*[@id="rs"]/div/table/tbody/tr[3]/th[3]/a').text
            except:
                i = '无相关数据'

            ai = fistCol[searchKyi]
            searchData = [ai, a, b, c, d, e, f, g, h, i]
            for m in range(len(searchData)):
                worksheet.write(searchKyi, m, searchData[m])
                print(searchData[m])
            browser.implicitly_wait(15)
            browser.find_element_by_id('kw').clear()

    # 保存
    workbook.close()
