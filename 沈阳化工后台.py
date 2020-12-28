# !/usr/bin/python3
# --coding:utf-8--
# @Author:ACHIEVE_DREAM
# @Time: 2020年12月17日22时
# @File: 沈阳化工后台.py
import os
from print_color import *
import requests
from lxml import etree
from yzm_recognize import yzm_result


class EducationSystem:
    # 主链
    base_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/'
    # 登录的网址
    login_url = f'{base_url}default2.aspx'
    # 验证码刷新链接
    yzm_url = f'{base_url}CheckCode.aspx'
    # 学生主页链接
    home_url = f'{base_url}xs_main.aspx?xh='
    # 储存所有的请求链接
    all_url = {}
    # 通用请求头
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,'
                  'image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'en,zh-CN;q=0.9,zh;q=0.8',
        'Host': '202.199.115.46',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                      ' Chrome/78.0.3904.108 '
                      'Safari/537.36'}
    user_id: str
    password: str
    name: str

    # 获取账号密码
    @staticmethod
    def getUserPwd():
        # name = input_c('账号: ', color1=fcyan)
        # pwd = input_c('密码: ', color1=fcyan)
        name = '1803030110'
        pwd = 'wulei123456'
        return name, pwd

    # 登录
    def login(self):
        self.user_id, self.password = self.getUserPwd()
        while True:
            erCode = yzm_result()
            viewState = self.getPageParm(self.login_url)
            # 登录时需要上传的表单数据,注意中文需要转成'gb2312'
            data = {'__VIEWSTATE': viewState,
                    'txtUserName': self.user_id,
                    'Textbox1': '',
                    'TextBox2': self.password,
                    'txtSecretCode': erCode,
                    'RadioButtonList1_2': '学生'.encode('gb2312'),
                    'Button1': '',
                    'lbLanguage': '',
                    'hidPdrs': '',
                    'hidsc': '',
                    }
            # 获取响应进行判断
            result = requests.post(self.login_url, data=data).content.decode('gb2312')
            if self.isRight(result):
                self.getInfos(result)
                print_c(f'欢迎{self.name},登录成功', color1=fgreen)
                os.remove('yzm.png')
                break

    # 判断是否登录成功
    def isRight(self, response):
        if "验证码不正确" in response:
            print_c('验证码不正确,正在重新识别...')
        elif "密码错误" in response:
            self.password = input_c('密码错误,请重新输入密码: ', color1=fcyan)
        elif "用户名不存在" in response:
            self.user_id = input_c('用户名不存在,请重新输入: ', color1=fcyan)
        else:
            return True
        return False

    # 获取信息
    def getInfos(self, response):
        html = etree.HTML(response)
        self.name = html.xpath("//div[@class='info']//span[@id='xhxm']/text()")[0]
        score_url = self.base_url + html.xpath("//ul/li[position()=6]/ul/li[position()=4]/a/@href")[0]
        self.all_url['score'] = score_url
        self.headers['Referer'] = self.home_url + self.user_id

    # 获取VIEWSTARE值
    def getPageParm(self, url):
        response = requests.get(url, headers=self.headers).content.decode('gb2312')
        html = etree.HTML(response)
        viewState = html.xpath("//form[@id='form1']/input[@name='__VIEWSTATE']/@value")
        if not viewState:
            viewState = html.xpath("//form[@id='Form1']/input[@name='__VIEWSTATE']/@value")
        return viewState[0]

    # 查询成绩
    def getScore(self):
        score_url = self.all_url.get('score')
        viewState = self.getPageParm(score_url)
        # 查询成绩需要上传的表单
        data = {
            '__VIEWSTATE': viewState,
            'ddlXN': '',
            'ddlXQ': '',
            'Button2': '在校学习成绩查询'.encode('gb2312')
        }
        result = requests.post(score_url, headers=self.headers, data=data).content.decode('gb2312')
        html = etree.HTML(result)
        tr_s = html.xpath('//*[@id="Datagrid1"]/tr')
        useful_index = [0, 1, 3, 6, 7, 8, 9, 10, 11, 12, 13]
        with open(self.name + '.csv', 'a', encoding='utf8') as f:
            for tr in tr_s:
                td_text = tr.xpath('td/text()')
                line_text = ''
                for i in useful_index:
                    score = 'none' if td_text[i] == '&nbsp;' else td_text[i]
                    line_text += score + ','
                line_text += '\n'
                f.write(line_text)
        print_c(f'已保存至{self.name}.csv', color1=fgreen)

    def run(self):
        self.login()
        self.getScore()


if __name__ == '__main__':
    educationSystem = EducationSystem()
    educationSystem.run()
