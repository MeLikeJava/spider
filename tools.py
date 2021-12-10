import re

import requests
from bs4 import BeautifulSoup


class DataResolve:
    def __init__(self, html):
        self.html = html

    # 提供外部访问的方法
    def resolve(self):
        return self.__resolvePicture()
        # self.__resolveScore()
        # self.__resolveEvaluations()
        # self.__resolveFilmDownLoad()

    # 解析图片
    def __resolvePicture(self):
        data = []# 数据列表
        a = 0 # 图片名
        if __name__ == '__main--':
            print("调用成功")
        findImgSrc = re.compile(r'<img.*src="(.*?)"', re.S)
        soup = BeautifulSoup(self.html, "html.parser")
        for item in soup.find_all('div', attrs={'class': 'item'}):
            item = str(item)
            # findall()返回的是一个列表,加[0]表示只返回索引为0的数据,不加则返回整个列表
            imgSrc = re.findall(findImgSrc, item)[0]
            data.append(imgSrc)
        for src in data:
            a += 1
            fileURL = 'C:\\Users\\snowball\\PycharmProjects\\spider\\img\\img' + str(a) + '.jpg'
            with open(fileURL, 'wb') as f:
                f.write(requests.get(src).content)
        return data

    # 解析评价信息
    def __resolveScore(self):
        return ''

    # 解析评价数
    def __resolveEvaluations(self):
        return ''

    # 解析电影下载地址
    def __resolveFilmDownLoad(self):
        return ''
