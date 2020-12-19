# !/usr/bin/python3
# --coding:utf-8--
# @Author:ACHIEVE_DREAM
# @Time: 2020年12月17日22时
# @File: 沈阳化工后台.py
import os

import requests
from lxml import etree


class EducationSystem:
    # 登录的网址
    login_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/default2.aspx'
    # 验证码刷新链接
    yzm_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/CheckCode.aspx'
    # 学生主页链接
    home_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/xs_main.aspx?xh='
    # 储存所有的请求链接
    all_url_data = {}
    # 最基本的链接
    base_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/'
    # 通用请求头
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,'
                  'image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'en,zh-CN;q=0.9,zh;q=0.8',
        'Connection': 'keep-alive',
        'Host': '202.199.115.46',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                      ' Chrome/78.0.3904.108 '
                      'Safari/537.36'}

    def __init__(self, user: str, password: str):
        self.user = user
        self.password = password

    def login(self):
        self.getVerificationCode()

    def getVerificationCode(self):
        while True:
            # 写入验证码并展示
            with open('yzm.jpg', 'wb')as f:
                f.write(requests.get(self.yzm_url, headers=self.headers).content)
            os.system('yzm.jpg')
            yzm = input('请输入验证码: ')
            os.system("taskkill /F /IM Microsoft.Photos.exe")
            break
        return yzm

    def run(self):
        pass


if __name__ == '__main__':
    name = '1803030110'
    pwd = 'wulei123456'
    educationSystem = EducationSystem(name, pwd)
    educationSystem.login()
