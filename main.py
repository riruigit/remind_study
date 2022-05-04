# 作者：海绵
# github地址：
# 作者mail: 1836426588@qq.com
# 介绍：前提是能登录微信网页版，要是不能登录网页版那算了。下面的当我在放屁算了。
# 先去青年大学习的网站，下载到未学习的名单，然后会把未学习的名单拉一个微信群，然后微信群里提醒各位青年大学习。
# 目前已经的缺点：代码无法删除群成员，只能增加（微信封了接口吧，不清楚）
# 建议点：看完了的同学自觉退群。
# 使用步骤：提前创建好群名是 “未大学习名单” 的微信群 ， 同时保存群到通讯录。


from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
import time
import os
from wxpy import *
from openpyxl import load_workbook
import warnings
import sys
import yaml

warnings.simplefilter("ignore")

option = ChromeOptions()
option.add_experimental_option("excludeSwitches", ["enable-automation"])
option.add_argument('--no-sandbox')
option.add_argument('--disable-dev-shm-usage')
option.add_argument('--start-maximized')
option.add_argument('--disable-gpu')
option.add_argument('--hide-scrollbars')
prefs = {"download.default_directory": os.getcwd()}
option.add_experimental_option("prefs", prefs)


def readfile(filepath):
    f = open(filepath, encoding="utf-8")
    yaml_reader = yaml.load(f.read(), Loader=yaml.FullLoader)
    f.close()
    return yaml_reader


def auto_down():
    bro = webdriver.Chrome(options=option)
    bro.get('https://tuan.12355.net/index.html')
    bro.find_element(By.XPATH, '//*[@id="userName"]').send_keys(config['userName'])
    bro.find_element(By.XPATH, '//*[@id="password"]').send_keys(config['password'])
    bro.find_element(By.XPATH, '//*[@id="login"]').click()
    time.sleep(2)
    bro.find_element(By.XPATH, '//*[@id="nav"]/div[8]/div[1]/div[1]').click()
    time.sleep(2)
    bro.switch_to.window(bro.window_handles[1])
    time.sleep(2)
    bro.find_element(By.XPATH, '//*[@id="app"]/div/div[1]/div[1]/div/ul/div[2]').click()
    time.sleep(2)
    bro.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/section/div/div[5]/div[3]/table/tbody/tr/td[5]').click()
    time.sleep(2)
    bro.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/section/div/div[3]/form/div[6]/div/button').click()
    time.sleep(2)
    bro.quit()


def auto_send():
    wb = load_workbook(filename='组织未参学名单【青年大学习】.xlsx')
    ws = wb['sheel1']
    to_send_user = []
    for i, name in enumerate(ws.values):
        if i >= 2:
            to_send_user.append(name[0])
    print("还没有参与青年大学习的名单", end='  ')
    print(to_send_user)
    bot = Bot(cache_path=True)
    wx_list = []
    for i, user in enumerate(to_send_user):
        find_user = bot.friends().search(str(user))
        if find_user is None or len(find_user) == 0:
            continue
        else:
            wx_user = find_user[0]
            wx_list.append(wx_user)
    print(wx_list)
    study_group = bot.groups().search(config['classroom'])[0]
    study_group.send("各位同学，我又来催青年大学习了")
    if len(wx_list) != 0:
        study_group.add_members(wx_list)
    study_group.send("搞快点，学习完了同学自觉退群")
    study_group.send("搞快点，学习完了同学自觉退群")
    study_group.send("搞快点，学习完了同学自觉退群")


config = readfile("config.yml")
auto_down()
auto_send()
os.remove('组织未参学名单【青年大学习】.xlsx')
