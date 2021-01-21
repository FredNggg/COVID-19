import math
# 爬取新浪新闻微博数据
import random

import requests
from bs4 import BeautifulSoup
import json
from openpyxl import workbook  # 写入Excel表所用
from openpyxl import load_workbook  # 读取Excel表所用

user_agent_list = [
    'Mozilla/5.0 (Windows NT 6.2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1464.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.16 Safari/537.36',
    'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.3319.102 Safari/537.36',
    'Mozilla/5.0 (X11; CrOS i686 3912.101.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.116 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1667.0 Safari/537.36',
    'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:17.0) Gecko/20100101 Firefox/17.0.6',
    'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1468.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2224.3 Safari/537.36',
    'Mozilla/5.0 (X11; CrOS i686 3912.101.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.116 Safari/537.36']


def crawl_url(page_num):
    print('Printing page' + str(page_num))
    url = 'https://weibo.cn/sinapapers?page=' + str(page_num)
    UserAgent = random.choice(user_agent_list)
    headers = {
        'User-Agent': UserAgent,
        'Cookie': 'SCF=Al4xjJm-QpE6K8stIcv58vYW_RukWN1n9SgtWTRkImUUxPYcG-wjrGwpc7YN4EnARmGsbDjzW3Khskz7RuvjwvU.; _T_WM=f44d1f2ca5957047135b25519bdcea12; SUB=_2A25y55VzDeRhGeNM71sV8y7Iyz-IHXVuKzs7rDV6PUJbkdAKLVbmkW1NTgFaZmGrQ6Tv0dyrY_mcCyI_XO3TmV3r; SSOLoginState=1608770851'}

    response = requests.get(url=url, headers=headers)
    page_text = response.text
    soup = BeautifulSoup(page_text, 'lxml')
    weibo_list = soup.find_all('div', class_='c')

    weibo_list = weibo_list[1:-2]
    for weibo in weibo_list:
        comment_url = weibo.find('a', class_='cc').get('href')
        with open('./weibo_commentURL.txt', 'a', encoding='utf-8') as fp:
            fp.write(comment_url + '\n')


# def whether_comment_is_block(soup): #判断微博评论是否被屏蔽


def crawl_text(weibo_url):
    global ws
    url = weibo_url
    UserAgent = random.choice(user_agent_list)
    headers = {
        'User-Agent': UserAgent,
        'Cookie': 'SCF=Al4xjJm-QpE6K8stIcv58vYW_RukWN1n9SgtWTRkImUUxPYcG-wjrGwpc7YN4EnARmGsbDjzW3Khskz7RuvjwvU.; _T_WM=f44d1f2ca5957047135b25519bdcea12; SUB=_2A25y55VzDeRhGeNM71sV8y7Iyz-IHXVuKzs7rDV6PUJbkdAKLVbmkW1NTgFaZmGrQ6Tv0dyrY_mcCyI_XO3TmV3r; SSOLoginState=1608770851'}

    response = requests.get(url=url, headers=headers)
    page_text = response.text
    soup = BeautifulSoup(page_text, 'lxml')
    c_text = soup.find_all('div', class_='c')
    comment_list = []
    temp = c_text[-2].text
    if temp == '''还没有人针对这条微博发表评论!''':  # 表示该微博评论被屏蔽
        comment_list.append(' ')
    else:
        comment_list = crawl_hot_comment(weibo_url)
    weibo_text = soup.find('div', class_='c', id="M_")

    date_text = weibo_text.find('span', class_='ct').text
    weibo_text = weibo_text.find('span', class_='ctt').text
    ws.append([date_text,weibo_url,weibo_text]+(comment_list))
    # print(weibo_text)

def crawl_hot_comment(weibo_url):
    UserAgent = random.choice(user_agent_list)
    headers = {
        'User-Agent': UserAgent,
        'Cookie': 'SCF=Al4xjJm-QpE6K8stIcv58vYW_RukWN1n9SgtWTRkImUUxPYcG-wjrGwpc7YN4EnARmGsbDjzW3Khskz7RuvjwvU.; _T_WM=f44d1f2ca5957047135b25519bdcea12; SUB=_2A25y55VzDeRhGeNM71sV8y7Iyz-IHXVuKzs7rDV6PUJbkdAKLVbmkW1NTgFaZmGrQ6Tv0dyrY_mcCyI_XO3TmV3r; SSOLoginState=1608770851'}

    hot_comment_list = []
    hot_url = 'https://weibo.cn/comment/hot/' + weibo_url[25:]
    response = requests.get(url=hot_url, headers=headers)
    page_text = response.text
    soup = BeautifulSoup(page_text, 'lxml')
    c_text = soup.find_all('div', class_='c')
    for i in range (2,5):
        if(i==len(c_text)):
            break
        hot_comment_text = c_text[i].find('span', class_='ctt').text
        hot_comment_list.append(hot_comment_text)
    return hot_comment_list
    # print(c_text)
# def crawl_comment():
#     i = 880
#     while (i < 1800):
#         crawl_url(i)
#         i = i + 1


if __name__ == "__main__":
    # crawl_comment()
    # crawl_text('https://weibo.cn/comment/ItytgEJEh?uid=2028810631&rl=0#cmtfrm')
    # crawl_hot_comment('https://weibo.cn/comment/IravqsEtQ?ckAll=1')
    url_list=[]
    file = open("./weibo_commentURL.txt")
    while 1:
        line = file.readline()[:-2]
        if not line:
            break
        url_list.append(line)
    file.close()
    wb = workbook.Workbook()  # 创建Excel对象
    ws = wb.active  # 获取当前正在操作的表对象
    # ws.append(['日期','链接', '正文','热评1','热评2','热评3'])
    count = 1
    for url in url_list:
        try:
            crawl_text(weibo_url=url)
            print('data'+str(count)+" saved successfully!")
            count = count +1
            if(count == 255):
                break
        except:
            print('wrong')
            wb.save('data'+str(count)+'.xlsx')
    wb.save('data.xlsx')