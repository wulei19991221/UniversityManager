# !/usr/bin/python3
# --coding:utf-8--
# @Author:吴磊
# @Time: 2019年12月06日17时
# @File: 我的考试系统信息查询.py
import msvcrt
import time
from threading import Thread
from queue import Queue
import requests
import re
from bs4 import BeautifulSoup
import xlwt
import os
import random


# 密码输入显示*号,方法来源于网络
def pwd_input(remain):
    print(remain, end='', flush=True)
    li = []
    while 1:
        ch = msvcrt.getch()
        # 回车
        if ch == b'\r':
            return b''.join(li).decode()  # 把list转换为字符串返回
        # 退格
        elif ch == b'\x08':
            if li:
                li.pop()
                msvcrt.putch(b'\b')
                msvcrt.putch(b' ')
                msvcrt.putch(b'\b')
        # Esc
        elif ch == b'\x1b':
            break
        else:
            li.append(ch)
            msvcrt.putch(b'*')
    return b''.join(li).decode()


#    随机评价（对老师友好点）
def get_comment():
    good = '优秀'.encode('gb2312')
    little_good = '良好'.encode('gb2312')
    mid_good = '中等'.encode('gb2312')
    comment_list = [good, little_good, mid_good]
    return random.choice(comment_list)


class EducationSystem(object):
    # 自动初始化
    def __init__(self):
        # 登录的网址
        self.login_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/default2.aspx'
        # 验证码刷新链接
        self.yzm_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/CheckCode.aspx'
        # 学生主页链接
        self.home_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/xs_main.aspx?xh='
        # 储存所有的请求链接
        self.all_url_data = {}
        # 最基本的链接
        self.base_url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/'
        # 通用请求头
        self.headers = {
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
        # 队列
        self.url_queue = Queue()
        # 根据网址获取正在评价的课程
        self.comment_index = {}
        # 待评价课程的数量
        self.comment_num = 0
        # 用户的基本信息
        self.all_data = {
            'userid': '',
            'password': '',
            'name': ''
        }
        # 课程评价的表单
        self.data = {
            '__EVENTTARGET': '',
            '__EVENTARGUMENT': '',
            'pjxx': '',
            'txt1': '',
            'TextBox1': '0',
            'Button1': '保  存'.encode('gb2312')
        }

    # 获取验证码
    def get_yzm(self):
        while True:
            # 写入验证码并展示
            with open('验证码.png', 'wb')as f:
                f.write(requests.get(self.yzm_url, headers=self.headers).content)
            os.system('验证码.png')
            yzm = input('请输入验证码: ')
            break
        return yzm

    # 登录,目的只是获取正确的账号和密码,方便以后的操作
    def login(self):
        if os.path.exists('account_password.txt'):
            content = open('account_password.txt', 'r', encoding='utf-8').read()
            name = re.findall(re.compile('姓名: (.*?),'), content)
            passwords = re.findall(re.compile('密码: (.*?),'), content)
            userids = re.findall(re.compile('学号; (.*?),'), content)
            is_new_man = input('你是否登录过这台设备y/n: ')
            if is_new_man == 'y':
                print("已登录账号:")
                for i, v in enumerate(name):
                    print(f"{i}:\t{v}")
                num = int(input("请输入对应序号: "))
                password = passwords[num]
                userid = userids[num]
            else:
                userid = input('请输入学号: ')
                password = pwd_input("请输入密码: ")
        else:
            userid = input('请输入学号: ')
            password = pwd_input("请输入密码: ")
        yzm = self.get_yzm()
        os.system("taskkill /F /IM Microsoft.Photos.exe")
        VIEWSTARE = self.get_VIEWSTARE(self.login_url)
        while True:
            # 登录时需要上传的表单数据,注意中文需要转成'gb2312'
            data = {'__VIEWSTATE': VIEWSTARE,
                    'txtUserName': userid,
                    'Textbox1': '',
                    'TextBox2': password,
                    'txtSecretCode': yzm,
                    'RadioButtonList1_2': '学生'.encode('gb2312'),
                    'Button1': '',
                    'lbLanguage': '',
                    'hidPdrs': '',
                    'hidsc': '',
                    }
            # 获取响应进行判断
            result = requests.post(self.login_url, data=data).content.decode('gb2312')
            if "验证码不正确" in result:
                print('验证码不正确,请重新输入!!!')
                yzm = self.get_yzm()
                os.system("taskkill /F /IM Microsoft.Photos.exe")
                continue
            elif "密码错误" in result:
                password = pwd_input('密码错误,请重新输入密码: ')
                yzm = self.get_yzm()
                continue
            elif "用户名不存在" in result:
                userid = input('用户名不存在,请重新输入: ')
                yzm = self.get_yzm()
                continue
            else:
                # 登录成功,将正确的信息写入数据包
                os.remove('./验证码.png')  # 删除缓存的验证码
                # 获取学生的姓名,学号,密码
                rule = re.compile('id="xhxm">(.*?)同学</span>')
                self.all_data['name'] = re.findall(rule, result)[0]
                self.all_data['userid'] = userid
                self.all_data["password"] = password
                # 暂时保存账号和密码
                self.writeinfo()
                # 所有名称及所有对应的请求网址
                rule2 = re.compile(r'<a href=".*?" target=\'zhuti\' onclick="GetMc\(\'.*?\'\);">(.*?)</a>')
                rule3 = re.compile(r'<a href="(.*?)" target=\'zhuti\' onclick="GetMc\(\'.*?\'\);">.*?</a>')
                result2 = re.findall(rule2, result)
                result3 = re.findall(rule3, result)
                for i in range(len(result2)):
                    self.all_url_data[result2[i]] = self.base_url + result3[i]
                # 待评价课程的数量
                rule4 = re.compile("<span class='down'> 教学质量评价</span>(.*?)<span class='down'> 信息维护")
                lessons = re.findall(rule4, result)[0]
                rule5 = re.compile('<li>(.*?)</li>')
                self.comment_num = len(re.findall(rule5, lessons))
                print('登录成功')
                break

    # 获取VIEWSTARE值
    def get_VIEWSTARE(self, url):
        rule = re.compile('"__VIEWSTATE" value="(.*?)" />')
        html = requests.get(url, headers=self.headers).content.decode('gb2312')
        VIEWSTARE = re.findall(rule, html)[0]
        return VIEWSTARE

    # 快捷评价
    def comment(self):
        # 课程评价的开始和结束位置,智能识别
        end_p = 0
        for k in self.all_url_data.keys():
            if k == '个人信息':
                break
            else:
                end_p += 1
        start_p = end_p - self.comment_num
        end_p -= 1
        # 评价课程和链接的索引
        all_class_name = list(self.all_url_data.keys())[int(start_p):int(end_p) + 1]
        all_comment_url = list(self.all_url_data.values())[int(start_p):int(end_p) + 1]
        if len(all_comment_url) == 0:
            print('你已经评价过啦')
            self.options()
        else:
            self.headers['Referer'] = self.home_url + self.all_data["userid"]
            li = 0
            # 将倒数第二及以前的链接放入多线程
            for comment1 in all_comment_url:
                # 入队列
                self.url_queue.put(comment1)
                # 创建执行的响应
                self.comment_index[comment1] = all_class_name[li]
                li += 1
            print('已开启多线程,评价系统速度取决于你的网速')
            # 创建多线程
            t_list = []
            for i in range(6):
                t = Thread(target=self.commentPlus(tips='正在评价: '))
                t_list.append(t)
                t.start()
            for i in t_list:
                i.join()
            # 最后一个课程需要完成提交操作
            last_comment_url = all_comment_url[-1]
            self.url_queue.put(last_comment_url)
            # 提交
            self.data['Button2'] = '提  交'.encode('gb2312')
            self.commentPlus(tips='正在提交...')
            print('评价完成')
            self.options()

    # 多线程评价
    def commentPlus(self, tips):
        while True:
            if self.url_queue.empty():
                break
            comment1 = self.url_queue.get()
            viewstare = self.get_VIEWSTARE(comment1)
            rule = re.compile('xkkh=(.*?)&')
            # 评价课程的编号
            pjkc = re.findall(rule, comment1)[0].encode('gb2312')
            content = requests.get(comment1, headers=self.headers).content.decode('gb2312')
            rule2 = re.compile('<select name="(.*?)" id=".*?">')
            rule3 = re.compile('<input name="(.*?)" type="text" id=".*?" style')
            # 增加两个变化的数据
            self.data['__VIEWSTATE'] = viewstare
            self.data['pjkc'] = pjkc
            for i in re.findall(rule2, content)[1:]:
                self.data[i] = get_comment()
            for i in re.findall(rule3, content):
                self.data[i] = ''
            if tips == '正在评价: ':
                print(tips + self.comment_index.get(comment1))
            else:
                print(tips)
            # 更新请求头
            self.headers['Referer'] = comment1
            requests.post(comment1, headers=self.headers, data=self.data)

    # 查询成绩
    def get_scores(self):
        # 获取查询成绩的网址
        stu_scores = self.all_url_data['成绩查询']
        # 获取分数界面的VIEWSTARE值,更新headers(确保请求必须是从主页发出）
        self.headers['Referer'] = self.home_url + self.all_data["userid"]
        viewstare = self.get_VIEWSTARE(stu_scores)
        # 查询成绩需要上传的表单
        data = {
            '__VIEWSTATE': viewstare,
            'ddlXN': '',
            'ddlXQ': '',
            'Button2': '在校学习成绩查询'.encode('gb2312')
        }
        result = requests.post(stu_scores, headers=self.headers, data=data).content.decode('gb2312')
        soup = BeautifulSoup(result, 'html.parser')
        # 将获取的数据分行读取
        lines = str(soup.find(id='Datagrid1')).splitlines()
        # 平均绩点
        average_grade = soup.find(id='pjxfjd').text.split('：')[-1]
        rule = re.compile("<td>(.*?)</td>")
        # 新建一个表格
        addsheet = xlwt.Workbook(encoding='utf-8')
        sheet = addsheet.add_sheet(self.all_data['name'] + '的成绩')
        # 需要的数据的索引
        index = [0, 3, 6, 7, 8, 9, 10, 11, 12, 14]
        # k:行 j:列
        k = 0
        # 由于前两列数据太长,所以需要加宽
        sheet.col(0).width = 70 * 50
        sheet.col(1).width = 120 * 50
        # 按行读取
        for line in lines:
            result = re.findall(rule, line)
            j = 0
            # 判断列表是否有数据
            if len(result) != 0:
                for i in index:
                    # 判断这个数据是否是空值
                    if result[i] == '\xa0':
                        result[i] = ''
                    sheet.write(k, j, result[i])
                    j += 1
                k += 1
        sheet.write(0, 10, '平均绩点')
        sheet.write(1, 10, average_grade)
        addsheet.save(self.all_data['name'] + '的成绩' + time.strftime("%Y-%m-%d") + '.xls')
        print(
            '你的成绩已经保存在:' + os.path.abspath(self.all_data['name'] + '的成绩' + time.strftime("%Y-%m-%d") + '.xls'))
        print('想要执行下一步,请关闭刚刚打开excel的软件(wps)')
        os.system(self.all_data['name'] + '的成绩' + time.strftime("%Y-%m-%d") + '.xls')
        self.options()

    # 选体育课
    def chooseclass(self):
        sports_url = self.all_url_data.get("选体育课")
        self.headers['Referer'] = self.home_url + self.all_data["userid"]
        result = requests.get(sports_url, headers=self.headers).content.decode('gb2312')
        soup = BeautifulSoup(result, 'html.parser')

    # 课程表查询
    def class_schedule(self):
        schedule_url = self.all_url_data['学生个人课表']
        self.headers['Referer'] = self.home_url + self.all_data["userid"]
        result = requests.get(schedule_url, headers=self.headers).content.decode('gb2312')
        soup = BeautifulSoup(result, 'html.parser')
        result = str(soup.find(id='Table1'))
        # 星期几
        rule = re.compile('<td align="Center" width="14%">(.*?)</td>')
        # 上课时间
        rule2 = re.compile('<td align="Center" rowspan="2".*?>(.*?)</td>')
        weekdays = re.findall(rule, result)
        schedule_time = re.findall(rule2, result)
        Monday_class = []
        Tuesday_class = []
        Wednesday_class = []
        Thursday_class = []
        Friday_class = []
        all_class = []
        data = {
            '1': '1,2',
            '2': '3,4',
            '3': '5,6',
            '4': '7,8',
            '5': '9,10'
        }
        for i in schedule_time:
            i = i.split('<br/>')
            schedule_time = ''
            for j in range(len(i)):
                schedule_time += i[j] + '\n'
            # 每节课的信息
            if '周一' in schedule_time:
                Monday_class.append(schedule_time)
            if '周二' in schedule_time:
                Tuesday_class.append(schedule_time)
            if '周三' in schedule_time:
                Wednesday_class.append(schedule_time)
            if '周四' in schedule_time:
                Thursday_class.append(schedule_time)
            if '周五' in schedule_time:
                Friday_class.append(schedule_time)
        all_class.append(Monday_class)
        all_class.append(Tuesday_class)
        all_class.append(Thursday_class)
        all_class.append(Wednesday_class)
        all_class.append(Friday_class)
        # 新建一个表格
        addsheet = xlwt.Workbook(encoding='utf-8')
        sheet = addsheet.add_sheet(self.all_data['name'] + '的课程表')
        style = xlwt.XFStyle()
        style.alignment.wrap = 1  # 设置自动换行
        for i in range(len(weekdays)):
            sheet.col(i).width = 256 * 20
        tall_style = xlwt.easyxf('font:height 720')
        # 第一行写星期
        for i in range(len(weekdays)):
            sheet.write(0, i, weekdays[i])

        for index, value in enumerate(all_class):
            for k, v in data.items():
                for i in value:
                    if v in i:
                        sheet.row(int(k)).set_style(tall_style)
                        sheet.write(int(k), index, i, style)
        addsheet.save(self.all_data['name'] + '的课程表.xls')
        print('你的课程表已经保存在:' + os.path.abspath(self.all_data['name']) + '的课程表.xls')
        print('想要执行下一步,请关闭刚刚打开excel的软件(wps)')
        os.system(self.all_data['name'] + '的课程表.xls')
        self.options()

    # 考试时间
    def exam_time(self):
        exam_url = self.all_url_data.get('学生考试查询')
        self.headers['Referer'] = self.home_url + self.all_data["userid"]
        content = requests.get(exam_url, headers=self.headers).content.decode('gb2312')
        soup = BeautifulSoup(content, 'html.parser')
        table = soup.table
        trs = table.find_all('tr')
        useful_index = [1, 2, 3, 4, 6]
        # 新建一个表格
        addsheet = xlwt.Workbook(encoding='utf-8')
        sheet = addsheet.add_sheet(self.all_data.get('name') + '的考试时间安排')
        sheet.col(2).width = 256 * 32
        sheet.col(3).width = 256 * 25
        sheet.col(0).width = 256 * 25
        row = 0
        for tr in trs:
            col = 0
            tds = tr.find_all('td')
            for index, td in enumerate(tds):
                if index in useful_index:
                    sheet.write(row, col, td.string)
                    col += 1
            row += 1
        addsheet.save(self.all_data.get('name') + '的考试时间安排.xls')
        print('你的课程表已经保存在:' + os.path.abspath(self.all_data.get('name') + '的考试时间安排.xls'))
        print('想要执行下一步,请关闭刚刚打开excel的软件(wps)')
        os.system(self.all_data.get('name') + '的考试时间安排.xls')
        self.options()

    #     首次进入函数
    def main(self):
        print('+' * 20 + '欢迎进入教务系统2.2,操作更简单,功能更强大 by 吴磊' + '+' * 20)
        print('首次进入需要登录')
        self.login()
        # 登录成功,提示用户进行下一步操作
        self.options()

    #   获取用户的操作
    def options(self):
        print('+' * 20 + "输入前面的序号,即可执行对应操作" + '+' * 20)
        print('1:\t成绩查询系统')
        print('2:\t网上选课系统')
        print('3:\t快捷评价系统')
        print('4:\t学生课表查询系统')
        print('5:\t考试时间查询系统')
        print('6:\t重新登录')
        print('7:\t退出操作系统')
        print('+' * 31 + '这是分割线' + '+' * 31)
        option = input()
        if option == '1':
            self.get_scores()
        elif option == '2':
            self.chooseclass()
        elif option == '3':
            self.comment()
        elif option == '4':
            self.class_schedule()
        elif option == '5':
            self.exam_time()
        elif option == '6':
            self.login()
            self.options()
        elif option == '7':
            exit()
        else:
            print('请执行正确的操作哟')
            self.options()

    # 储存个人信息
    def writeinfo(self):
        if os.path.exists('account_password.txt'):
            content = open('account_password.txt', 'r', encoding='utf-8').read()
            if self.all_data['name'] in content:
                pass
            else:
                open('account_password.txt', 'a', encoding='utf-8').write("姓名: " + self.all_data['name'] +
                                                                          ',\n学号: ' + self.all_data['userid'] +
                                                                          ',\n密码: ' + self.all_data[
                                                                              'password'] + ',\n')
        else:
            open('account_password.txt', 'w', encoding='utf-8').write("姓名: " + self.all_data['name'] +
                                                                      ',\n学号: ' + self.all_data['userid'] +
                                                                      ',\n密码: ' + self.all_data[
                                                                          'password'] + ',\n')


if __name__ == '__main__':
    test = EducationSystem()
    test.main()
