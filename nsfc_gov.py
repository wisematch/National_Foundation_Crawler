# Adopt from https://blog.csdn.net/qq_43814415/article/details/122271025

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
import time as t
import re
from bs4 import BeautifulSoup
import xlrd
import xlwt
import os
import random
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
# service = Service('/home/wei/Chrome/chromedriver-linux64/chromedriver')
options = Options()
options.add_argument("--mute-audio")  # 将浏览器静音
# chrome_options.add_experimental_option("detach", True)  # 当程序结束时，浏览器不会关闭
options.add_argument("--headless")
# options.add_argument("--disable-gpu")


# 先进入科学基金网
# driver = webdriver.Chrome(options=chrome_options, service=service)

driver = webdriver.Edge(options=options)

driver.maximize_window()
driver.get('https://kd.nsfc.gov.cn/finalProjectInit?advanced=true')     # 进入高级搜索页面


def analyse_most(html):
    """进行网页解析"""
    soup = BeautifulSoup(html, 'html.parser')  # 解析器：html.parser
    items=soup.select('html body div.hom div#wrapper_body div#main.wp div#mainCt div#resultLst.resultLst div.item')#一页的项目
    print(items)

    titles = []
    authors = []
    inses = []
    sorts = []
    nums = []
    years = []
    mons = []
    keys = []
    # 使用路径
    title = soup.select(
        'html body div.hom div#wrapper_body div#main.wp div#mainCt div#resultLst.resultLst div.item p.t a')  # 项目名
    author = soup.select(
        'html body div.hom div#wrapper_body div#main.wp div#mainCt div#resultLst.resultLst div.item div.d p.ico span.author i')  # 负责人
    # 使用选择器
    ins = []
    sort = []
    num = []
    year = []
    mon = []
    key = []
    for i in range(1, 11):
        ins_1 = soup.select('div.item:nth-child(' + str(
            i) + ') > div:nth-child(2) > p:nth-child(1) > span:nth-child(2) > i:nth-child(1)')  # 机构
        sort_1 = soup.select(
            'div.item:nth-child(' + str(i) + ') > div:nth-child(2) > p:nth-child(1) > i:nth-child(3)')  # 类型
        num_1 = soup.select(
            'div.item:nth-child(' + str(i) + ') > div:nth-child(2) > p:nth-child(1) > b:nth-child(4)')  # 批准号
        year_1 = soup.select('div.item:nth-child(' + str(
            i) + ') > div:nth-child(2) > p:nth-child(1) > span:nth-child(5) > b:nth-child(1)')  # 立项年份
        mon_1 = soup.select('div.item:nth-child(' + str(
            i) + ') > div:nth-child(2) > p:nth-child(2) > span:nth-child(1) > b:nth-child(1)')  # 资助金额
        key_1 = soup.select('div.item:nth-child(' + str(
            i) + ') > div:nth-child(2) > p:nth-child(2) > span:nth-child(2) > i:nth-child(1)')  # 关键词
        ins.append(ins_1)
        sort.append(sort_1)
        num.append(num_1)
        year.append(year_1)
        mon.append(mon_1)
        key.append(key_1)
    clear(title, titles)
    clear(author, authors)
    clear_l(ins, inses)
    clear_l(sort, sorts)
    clear_l(num, nums)
    clear_l(year, years)
    clear_l(mon, mons)
    clear_l(key, keys)

    page = []  # 本页的所有信息
    for i in range(len(titles)):
        page.append(titles[i:i + 1] + authors[i:i + 1] + inses[i:i + 1] + sorts[i:i + 1] + nums[i:i + 1] + years[
                                                                                                           i:i + 1] + mons[
                                                                                                                      i:i + 1] + keys[
                                                                                                                                 i:i + 1])
    return page


def analyse_end(html):
    """进行最后一页的网页解析"""
    soup = BeautifulSoup(html, 'html.parser')  # 解析器：html.parser
    # items=soup.select('html body div.hom div#wrapper_body div#main.wp div#mainCt div#resultLst.resultLst div.item')#一页的项目
    # print(items)
    titles = []
    authors = []
    inses = []
    sorts = []
    nums = []
    years = []
    mons = []
    keys = []
    # 使用路径
    title = soup.select(
        'html body div.hom div#wrapper_body div#main.wp div#mainCt div#resultLst.resultLst div.item p.t a')  # 项目名
    author = soup.select(
        'html body div.hom div#wrapper_body div#main.wp div#mainCt div#resultLst.resultLst div.item div.d p.ico span.author i')  # 负责人
    # 使用选择器
    ins = []
    sort = []
    num = []
    year = []
    mon = []
    key = []
    for i in range(1, 3):
        ins_1 = soup.select('div.item:nth-child(' + str(
            i) + ') > div:nth-child(2) > p:nth-child(1) > span:nth-child(2) > i:nth-child(1)')  # 机构
        sort_1 = soup.select(
            'div.item:nth-child(' + str(i) + ') > div:nth-child(2) > p:nth-child(1) > i:nth-child(3)')  # 类型
        num_1 = soup.select(
            'div.item:nth-child(' + str(i) + ') > div:nth-child(2) > p:nth-child(1) > b:nth-child(4)')  # 批准号
        year_1 = soup.select('div.item:nth-child(' + str(
            i) + ') > div:nth-child(2) > p:nth-child(1) > span:nth-child(5) > b:nth-child(1)')  # 立项年份
        mon_1 = soup.select('div.item:nth-child(' + str(
            i) + ') > div:nth-child(2) > p:nth-child(2) > span:nth-child(1) > b:nth-child(1)')  # 资助金额
        key_1 = soup.select('div.item:nth-child(' + str(
            i) + ') > div:nth-child(2) > p:nth-child(2) > span:nth-child(2) > i:nth-child(1)')  # 关键词
        ins.append(ins_1)
        sort.append(sort_1)
        num.append(num_1)
        year.append(year_1)
        mon.append(mon_1)
        key.append(key_1)
    clear(title, titles)
    clear(author, authors)
    clear_l(ins, inses)
    clear_l(sort, sorts)
    clear_l(num, nums)
    clear_l(year, years)
    clear_l(mon, mons)
    clear_l(key, keys)

    page = []  # 本页的所有信息
    for i in range(len(titles)):
        page.append(titles[i:i + 1] + authors[i:i + 1] + inses[i:i + 1] + sorts[i:i + 1] + nums[i:i + 1] + years[
                                                                                                           i:i + 1] + mons[
                                                                                                                      i:i + 1] + keys[
                                                                                                                                 i:i + 1])
    return page


def get_one(year):
    """返回搜索第一页源码"""
    # 选择年份
    if year == 2020:
        driver.find_element(By.XPATH, '/html/body/div[3]/div[1]/div/div[2]/table[2]/tbody/tr[1]/td[1]').click()   # 2020
    elif year == 2021:
        driver.find_element(By.XPATH, '/html/body/div[3]/div[1]/div/div[2]/table[2]/tbody/tr[1]/td[2]').click()   # 2021
    elif year == 2022:
        driver.find_element(By.XPATH, '/html/body/div[3]/div[1]/div/div[2]/table[2]/tbody/tr[1]/td[3]').click()   # 2022

    driver.find_element(
        '/html/body/div[2]/div[4]/div/div[1]/div/div[2]/div[1]/form/div[1]/div[2]/div[2]/span[2]/span[1]/span/span[1]').click()  # 直接点击
    t.sleep(2)
    driver.find_element('/html/body/dialog/bd/div[1]/div[1]/ul/li[8]').click()  # 选中管理科学部
    t.sleep(1)
    driver.find_element('/html/body/dialog/bd/div[1]/div[1]/div/ul/li[2]/label').click()  # 选中工商管理
    t.sleep(1)
    driver.find_element('/html/body/dialog/bd/div[1]/div[1]/div/div/ul/li[6]/label').click()  # 选中会计
    t.sleep(1)
    driver.find_element('/html/body/dialog/bd/div[2]/button').click()  # 确定
    t.sleep(3)
    driver.find_element(
        '/html/body/div[2]/div[4]/div/div[1]/div/div[2]/div[1]/form/div[1]/div[1]/div[3]/select[1]').click()  # 点击起始年份
    driver.find_element(
        '/html/body/div[2]/div[4]/div/div[1]/div/div[2]/div[1]/form/div[1]/div[1]/div[3]/select[1]/option[24]').click()  # 选中2000年
    t.sleep(3)
    driver.find_element(
        '/html/body/div[2]/div[4]/div/div[1]/div/div[2]/div[1]/form/div[1]/div[1]/div[3]/select[2]')  # 点击截止年份
    driver.find_element(
        '/html/body/div[2]/div[4]/div/div[1]/div/div[2]/div[1]/form/div[1]/div[1]/div[3]/select[2]/option[4]').click()  # 选中2020年
    t.sleep(1)
    driver.find_element(
        '/html/body/div[2]/div[4]/div/div[1]/div/div[2]/div[1]/form/div[2]/button').click()  # 点击搜索
    # 成功进入搜索结果页面
    html = driver.page_source  # 页面源码
    return html


def get_one_normal():
    # 面上项目
    get_one()  # 搜索首页
    driver.find_element_by_xpath('//*[@id="category"]').click()  # 进入类型
    t.sleep(1)
    driver.find_element_by_xpath('//*[@id="面上项目"]').click()  # 选定面上项目
    t.sleep(1)
    driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[2]/button').click()  # 点击筛选
    html = driver.page_source
    return html


# def get_one_qing():
#     # 青年科学基金项目，共24页
#     get_one()  # 搜索首页
#     driver.find_element_by_xpath('//*[@id="category"]').click()  # 进入类型
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="青年科学基金项目"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[2]/button').click()  # 点击筛选
#     html = driver.page_source
#     return html
#
#
# def get_one_others():
#     # 杂项，共9页
#     get_one()  # 搜索首页
#     driver.find_element_by_xpath('//*[@id="category"]').click()  # 进入类型
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="地区科学基金项目"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="专项基金项目"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="国际(地区)合作与交流项目"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="重点项目"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="应急管理项目"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="优秀青年科学基金项目"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="海外及港澳学者合作研究基金"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="重大项目"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="国家杰出青年科学基金"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('//*[@id="重大研究计划"]').click()  # 选定面上项目
#     t.sleep(1)
#     driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[2]/button').click()  # 点击筛选
#     html = driver.page_source
#     return html


def get_fore():
    """前6页的翻页，并返回源码"""
    driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[13]/a').click()  # 点击下一页
    html = driver.page_source
    return html


def get_fore_others():
    """前6页的翻页，并返回源码"""
    driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[11]/a').click()  # 点击下一页
    html = driver.page_source
    return html


def get_mid():
    """中间的翻页，并返回源码"""
    driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[15]/a ').click()  # 点击下一页
    html = driver.page_source
    return html


def get_after():
    """最后的翻页，并返回源码"""
    driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[14]/a ').click()  # 点击下一页
    html = driver.page_source
    return html


def clear(old_list, new_list):
    """用于清洗出纯文本"""
    for i in old_list:
        n = (i.text).strip()
        n = n.replace('\n', ' ')
        new_list.append(n)
    return new_list


def clear_l(old_list, new_list):
    """用于清洗出resultset的纯文本"""
    for i in old_list:
        i = i[0]
        n = (i.text).strip()
        n = n.replace('\n', ' ')
        new_list.append(n)
    return new_list


def save_afile(alls, count):
    os.chdir(r'E:\基金')
    """将一页的基金数据保存在一个excel"""
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    sheet1.write(0, 0, '项目名称')
    sheet1.write(0, 1, '负责人')
    sheet1.write(0, 2, '申请单位')
    sheet1.write(0, 3, '研究类型')
    sheet1.write(0, 4, '项目批准号')
    sheet1.write(0, 5, '批准年度')
    sheet1.write(0, 6, '资助金额')
    sheet1.write(0, 7, '关键词')
    i = 1
    for data in alls:  # 遍历每一行
        for j in range(len(data)):  # 取每一单元格
            sheet1.write(i, j, data[j])  # 写入单元格
        i = i + 1  # 往下一行
    f.save(str(count) + '.xls')
    print(str(count) + '保存成功！')


def wait():
    """返回随机等待时间"""
    s = t.sleep(random.randint(2, 9))
    return s


def wait_l():
    s = t.sleep(random.randint(1, 5))
    return s


def goto_10_mian():
    # 跳转到10页
    get_one_mian()
    wait_l()
    driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[9]/a').click()  # 转到第八页
    # for i in range(5):#转到第23页
    wait_l()
    # driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[11]/a').click()
    driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[10]/a').click()  # 转到第10页


# def goto_13_qing():
#     # 转到第13页
#     get_one_qing()
#     wait_l()
#     driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[9]/a').click()  # 转到第八页
#     wait_l()
#     driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[11]/a').click()  # 十一页
#     wait_l()
#     driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[10]/a').click()  # 十三页


def goto_20_mian():
    # 转到第20页
    get_one_mian()
    wait_l()
    driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[9]/a').click()  # 转到第八页
    for i in range(4):  # 转到第20页
        wait_l()
        driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[11]/a').click()


# def goto_22_qing():
#     # 转到第22页
#     get_one_qing()
#     wait_l()
#     driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[9]/a').click()  # 转到第八页
#     for i in range(4):  # 转到第20页
#         wait_l()
#         driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[11]/a').click()
#     driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[11]/a').click()  # 22


if __name__ == '__main__':
    # 进入面上项目第10页，因为已经爬到了该页数据
    # 进入面上项目第20页，因为已经爬到了该页数据
    # 进入青年基金第13页，因为已经爬到了该页数据
    for i in range(1, 10):  # 循环24次
        if i == 1:
            save_afile(analyse_most(get_one_normal()), i)
        elif 1 < i <= 8:
            save_afile(analyse_most(get_fore_others()), i)
        else:
            save_afile(analyse_end(get_fore_others()), i)

"""
        if i==1:
            save_afile(analyse_most(get_one_mian()),i)
        elif 1<i<=7:
            wait()
            save_afile(analyse_most(get_fore()),i)
            """
# driver.find_element_by_xpath('/html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[13]/a').click()#点击下一页
# /html/body/div[3]/div[4]/div/div[3]/div[2]/p/span[15]/a   7之后（最多一次弄七个，需要等待五分钟）

