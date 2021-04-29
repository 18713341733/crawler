# -*- coding:utf-8 -*-
# threadingdatasougou.py
from selenium import webdriver
import time
import xlrd
import xlsxwriter
import time
import os
import threading
import os
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep,ctime
import time
import xlrd
import xlsxwriter
import time
import openpyxl



url = 'https://www.sogou.com/web?query=%E5%BC%80%E5%A7%8B%E6%90%9C%E7%B4%A2&_ast=1616067877&_asf=www.sogou.com&w=01029901&p=40040108&dp=1&cid=&s_from=result_up&sut=3115&sst0=1616067882520&lkt=0%2C0%2C0&sugsuv=1602556158761669&sugtime=1616067882520'


def crawler(query_path):
    # 传参 query_path ，文件的绝对位置
    wenjianlist = os.listdir(query_path)
    # wenjianlist 读取query_path文件夹下所有的xlsx文件，并生成一个list
    # 如 ['备孕1.xlsx', '备孕2.xlsx']

    # 给生成的新的文件，命名
    namelist = []
    # 被读取的xlsx文件的绝对位置
    PathList = []
    for single in wenjianlist:
        newsingle = single.split('.')[0] + '的搜索结果.xlsx'
        namelist.append(newsingle)
        # 要读取的每个文件的位置
        PathList.append(query_path + "\\" + single)

    for namei in range(len(wenjianlist)):
        # 以xlsx文件为单位，开始循环
        time.sleep(5)


        ##################################################
        # 创建一个workbook 设置编码
        workbook = xlsxwriter.Workbook(namelist[namei])

        ###############################################
        # 初始化浏览器

        browser = webdriver.Chrome(r'D:\python3\Lib\chromedriver.exe')
        browser.get(url)
        browser.maximize_window()





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
                browser.implicitly_wait(6)



                try:
                    # 清理输入框
                    browser.find_element_by_id('upquery').click()
                    time.sleep(1)
                    browser.find_element_by_id('upquery').clear()
                    time.sleep(1)
                    browser.find_element_by_id('upquery').send_keys(fistCol[searchKyi])

                except:
                    time.sleep(2)
                    # 清理输入框
                    browser.find_element_by_id('upquery').click()
                    time.sleep(1)
                    browser.find_element_by_id('upquery').clear()
                    time.sleep(1)
                    browser.find_element_by_id('upquery').send_keys(fistCol[searchKyi])

                browser.implicitly_wait(6)
                # 点击回车，进行搜索
                # 回车进行搜索
                browser.find_element_by_id('upquery').send_keys(Keys.ENTER)

                time.sleep(2)
                searchlook = browser.find_elements_by_class_name('r-sech')
                searcha = searchlook[1].text
                searchlist = searcha.split()[1:]
                browser.implicitly_wait(6)

                ai = fistCol[searchKyi]
                searchlist.insert(0,ai)

                searchData = searchlist
                for m in range(len(searchData)):
                    worksheet.write(searchKyi, m, searchData[m])
                    print(searchData[m])
                time.sleep(1)


        # 保存
        workbook.close()
        browser.quit()


# 读取 query2文件夹下所有的文件，list
query_path1 = r'D:\babytree\codetest\zuoye\pachongsougou\query'
query_path2 = r'D:\babytree\codetest\zuoye\pachongsougou\query1'
# query_path3 = r'D:\babytree\codetest\zuoye\pachong\query4'
# query_path4 = r'D:\babytree\codetest\zuoye\pachong\query5'
# query_path5 = r'D:\babytree\codetest\zuoye\pachong\query6'
# query_path6 = r'D:\babytree\codetest\zuoye\pachong\query7'
threads = []
t1 = threading.Thread(target=crawler, args=(query_path1,))
threads.append(t1)
t2 = threading.Thread(target=crawler, args=(query_path2,))
threads.append(t2)
# t3 = threading.Thread(target=crawler, args=(query_path3,))
# threads.append(t3)
# t4 = threading.Thread(target=crawler, args=(query_path4,))
# threads.append(t4)
# t5 = threading.Thread(target=crawler,args=(query_path5,))
# threads.append(t5)
# t6 = threading.Thread(target=crawler,args=(query_path6,))
# threads.append(t6)
if __name__ == '__main__':
    for t in threads:
        t.setDaemon(True)
        t.start()

    for t in threads:
        t.join()
