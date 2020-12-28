# !/usr/bin/python3
# --coding:utf-8--
# @Author:ACHIEVE_DREAM
# @Time: 2020年12月28日14时
# @File: test.py
from aip import AipOcr
import requests
from API_INFOS import BaiduApi
import cv2
from PIL import Image


def yzm_result():
    client = AipOcr(BaiduApi.APP_ID, BaiduApi.API_KEY, BaiduApi.SECRET_KEY)
    url = 'http://202.199.115.46/(n2woc445sw0mzi55zdl1fojl)/CheckCode.aspx'
    while True:
        with open('yzm.png', 'wb') as f:
            content = requests.get(url).content
            f.write(content)
        Image.open('yzm.png').convert('L').save('yzm.png', format='PNG')
        img = cv2.imread('yzm.png')
        ret, img = cv2.threshold(img, 100, 255, cv2.THRESH_BINARY)
        cv2.imwrite('yzm.png', img)
        try:
            result = client.basicAccurate(open('yzm.png', 'rb').read())['words_result'][0]['words'].strip()
            if len(result) == 4:
                return result
        except (IndexError, KeyError):
            pass
