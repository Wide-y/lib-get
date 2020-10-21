# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.webdriver.support.select import Select
from bs4 import BeautifulSoup
import time
import xlwt
import xlrd

chrome_options = webdriver.ChromeOptions()
browser = webdriver.Chrome(chrome_options=chrome_options)


select_date_header = '2020-10-'  # 预约月份
select_date_start = 22  # 开始日期
select_date_end = 32  # 结束日期
select_time_start = '13:00' #开始时间
select_time_stop = '17:00'  #结束时间
add_num = 20  # 加入预约个数

groupname = '李铎大佬带飞组'
topic = '李铎大佬手把手辅导菜鸡'


def login():
    try:
        print('start login ...')
        login_url = 'https://jaccount.sjtu.edu.cn/jaccount/jalogin?sid=jalibtest04423&returl=%43%46%56%58%6E%74%6E%79%4F%6A%76%34%6C%50%5A%50%69%69%43%71%59%32%49%75%66%45%62%79%65%35%6D%4E%43%41%50%59%76%76%65%67%66%6A%36%6E%68%6A%55%50%42%73%30%6D%56%33%61%41%44%30%4B%2F%6C%42%72%67%46%63%55%58%53%32%4D%42%70%6B%6B%6C&se=%43%47%51%50%45%6C%69%30%5A%74%52%36%57%55%55%2F%50%4E%7A%4C%4B%37%79%65%66%56%68%2F%2F%74%32%51%38%77%3D%3D'
        home_url = 'http://studyroom.lib.sjtu.edu.cn/index.asp'
        browser.get(login_url)
        while(browser.current_url != home_url):  # 监测url变化确定登陆是否完成
            print('waiting...')
            time.sleep(5)
    except:
        print('error')
    finally:
        print('login done')


def begin(n):
    try:
        book_url = 'http://studyroom.lib.sjtu.edu.cn/apply.asp'
        browser.get(book_url)
        time.sleep(2)

        browser.find_element_by_name('date_s').clear()
        browser.find_element_by_name('date_s').send_keys(
            select_date_header+str(n))
        Select(browser.find_element_by_name('tstart')).select_by_value(select_time_start)
        Select(browser.find_element_by_name('tend')).select_by_value(select_time_stop)
        browser.find_element_by_class_name('fa-btn').click()
        sleep(1)
        # print(select_date_header+str(n))
    except:
        print('fetch error')
    finally:
        print('fetch done')


def choose_room():
    try:
        browser.find_element_by_link_text('预约').click()
        time.sleep(1)
        browser.find_element_by_name('groupname').send_keys(groupname)
        browser.find_element_by_name('topic').send_keys(topic)
        browser.find_element_by_name('detail').send_keys(topic)
        browser.find_element_by_name('attendcount').send_keys('3')
        browser.find_element_by_name('B1').click()
        time.sleep(1)
        a1 = browser.switch_to.alert
        time.sleep(1)
        a1.accept()
    except:
        print('choose room error')
    finally:
        print('choose done')


def get_pwd():
    try:
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('pwd')

        pwd_url = 'http://studyroom.lib.sjtu.edu.cn/user_reserve_list.asp'
        # pwd_url = 'file:///D:/cproject/python/lib_get/1.html'
        browser.get(pwd_url)
        time.sleep(5)

        soup = BeautifulSoup(browser.page_source, 'lxml')
        body = soup.find('body')
        trs = body.find_all('tr')[0:add_num+1]
        # print(trs)

        for i in range(1, add_num+1):
            tds = trs[i].find_all('td')
            print(tds[1])
            worksheet.write(i-1, 0, tds[1].getText())
            text = tds[10].getText()
            text = text[-7:]
            text = text[0:6]
            worksheet.write(i-1, 1, text)

        workbook.save('pwd.xls')
    except:
        print('getpwd error')
    finally:
        print('getpwd done')


def join(id, pwd):
    join_url = 'http://studyroom.lib.sjtu.edu.cn/reserve_plus.asp'
    browser.get(join_url)

    browser.find_element_by_name('applicationid').send_keys(id)
    browser.find_element_by_name('B1').click()

    time.sleep(1)

    browser.find_element_by_name('password').send_keys(pwd)
    browser.find_elements_by_name('B1')[1].click()
    time.sleep(1)
    a1 = browser.switch_to.alert
    time.sleep(1)
    a1.accept()
    


def join_pack():
    book = xlrd.open_workbook('pwd.xls')
    sheet1 = book.sheets()[0]
    for i in range(0, add_num):
        id = str(sheet1.cell_value(i, 0))[0:6]
        pwd = str(sheet1.cell_value(i, 1))[0:6]
        print(id, pwd)
        join(id, pwd)


def get():
    for i in range(select_date_start, select_date_end,1):
        begin(i)
        time.sleep(2)
        choose_room()
        time.sleep(2)
# def test():
#     s = '等待加入！(密码:321321)'
#     s = s[-7:]
#     print(s)
#     s = s[0:6]
#     print(s)


if __name__ == '__main__':
    login() #登陆

#以下功能自选其一取消注释即可
    #1. 自动发起预约
    # get()

    #2. 获取申请密码用
    # get_pwd()

    #3. 自动加入申请
    # join_pack()

    browser.close()

    # test()
    exit()
