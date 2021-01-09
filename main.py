# !/usr/bin/python3
# --coding:utf-8--
# @Author:ACHIEVE_DREAM
# @Time: 2020年12月30日14时
# @File: main.py
from SYHG_System import *


def init():
    system = EducationSystem()
    system.login()
    return system


def welcome():
    all_func = ['查分', '快速评价', '切换账号', '退出']
    print_c('欢迎使用教务处脚本版by ACHIEVE_DREAM'.center(LINE_WIDTH, FILL_CHAR), color1=fpurple)
    print_c('正在初始化, 需要先登录....', color1=fyellow)
    system: EducationSystem = init()
    print_c('(★ ω ★)我是有下线的(★ ω ★)'.center(LINE_WIDTH, FILL_CHAR), color1=fpurple)
    for i, func in enumerate(all_func, 1):
        print_c(f'❀ {i}:\t{func}', color1=fyellow)
    print_c('温馨提示: 输入对应序号,进行使用哦ヾ(≧▽≦*)o', color1=fyellow)
    print_c('(★ ω ★)我是有下线的(★ ω ★)'.center(LINE_WIDTH, FILL_CHAR), color1=fpurple)
    while True:
        try:
            opt = int(input_c('那么你想使用哪个功能呢?请告诉我吧♪(^∇^*): ', color1=fyellow))
            # opt = 2
            if opt == 1:
                system.getScore()
            elif opt == 2:
                system.quickComment()
            elif opt == 3:
                system.login()
            elif opt == 4:
                print_c('再见了, 我心爱的主人~~>_<~~', color1=fblue)
                exit()
            else:
                print_c('你输入的不正确哦,仔细看看吧::>_<::')
        except ValueError:
            print_c('你输入的不正确哦,仔细看看吧::>_<::')


if __name__ == '__main__':
    welcome()
