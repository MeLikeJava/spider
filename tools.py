import random
import re

import requests
from bs4 import BeautifulSoup
import xlwt
import xlrd
from xlutils.copy import copy
import time

from pyecharts.charts import WordCloud
from pyecharts import options as opts
from pyecharts.charts import TreeMap
from pyecharts.globals import SymbolType
from pyecharts.charts import Scatter

'''
    数据解析类
'''


class DataResolve:
    pictureData = {}  # 电影海报字典
    filmScore = {}  # 电影评分字典
    numberOfReviews = {}  # 电影评论数字典

    def __init__(self, html):
        self.html = html

        try:
            xlrd.open_workbook(".\\data\\filmData.xls")
        except:
            # 初始化本地存储excel表格
            print("filmData.xls文件不存在,创建filmData.xls文件")
            workbook = xlwt.Workbook()  # 创建工作簿对象
            workbook.add_sheet("sheet1")  # 创建工作表对象
            workbook.save(".\\data\\filmData.xls")
        try:
            open(".\\data\\logs.txt")
        except:
            print("创建logs.txt日志文件")
            open(".\\data\\logs.txt", 'w')  # 初始化日志文件

    '''
        提供外部访问方法
    '''

    def resolve(self):
        Utils.write_log()  # 日志记录
        picture = self.__resolvePicture()
        score = self.__resolveScore()
        film = self.__resolveEvaluations()
        # self.__resolveFilmDownLoad()
        return '解析结束'

    '''
        解析图片
    '''

    def __resolvePicture(self):
        if __name__ == '__main__':
            print("调用成功")
        findImgSrc = re.compile(r'<img.*src="(.*?)"', re.S)
        findFilmName = re.compile(r'<span class="title">(.*?)</span>')
        soup = BeautifulSoup(self.html, "html.parser")
        for item in soup.find_all('div', attrs={'class': 'item'}):
            item = str(item)
            filmName = re.findall(pattern=findFilmName, string=item)[0]
            # findall()返回的是一个列表,加[0]表示只返回索引为0的数据,不加则返回整个列表
            imgSrc = re.findall(findImgSrc, item)[0]
            self.pictureData[filmName] = imgSrc
        for item in self.pictureData.items():  # 遍历字典中的所有项

            imageURL = '.\\img\\' + item[0] + '.jpg'  # item[0]获取电影名称
            with open(imageURL, 'wb') as f:  # 循环写入文件
                f.write(requests.get(item[1]).content)  # item[1]电影海报的url地址
        time.sleep(0.1)
        print("图片解析成功,图片保存至img文件夹中")
        return self.pictureData

    '''
        解析评分
    '''

    def __resolveScore(self):
        findScore = re.compile(r'<span class="rating_num" property="v:average">(.*?)</span>', re.S)
        findFilmName = re.compile(r'<span class="title">(.*?)</span>')
        soup = BeautifulSoup(self.html, "html.parser")
        for item in soup.find_all(name='div', attrs={'class': 'item'}):
            item = str(item)
            score = re.findall(pattern=findScore, string=item)[0]
            filmName = re.findall(pattern=findFilmName, string=item)[0]
            self.filmScore[str(filmName)] = float(score)
        time.sleep(0.1)
        print("影片评分解析成功")
        return self.filmScore


    '''
        解析评价数
    '''

    def __resolveEvaluations(self):
        findNumberOfReviews = re.compile(r'<span>(\d.+)</span>')  # 匹配电影评价数
        findFilmName = re.compile(r'<span class="title">(.*?)</span>')  # 匹配电影名称
        soup = BeautifulSoup(self.html, "html.parser")
        for item in soup.find_all(name='div', attrs={'class': 'item'}):
            item = str(item)
            numberOfReviews = re.findall(pattern=findNumberOfReviews, string=item)[0]
            numberOfReviews = re.findall(pattern=r'\d+', string=numberOfReviews)[0]  # 剔除中文字符串
            filmName = re.findall(pattern=findFilmName, string=item)[0]
            self.numberOfReviews[str(filmName)] = int(numberOfReviews)
        time.sleep(0.1)
        print("影片评论条数解析成功")
        return self.numberOfReviews

    '''
        解析电影下载地址
    '''

    def __resolveFilmDownLoad(self):
        return ''

    '''
        (弃用)将数据写入excel
    '''

    def saveData(self, path):
        workbook = xlwt.Workbook(encoding='ascii')  # 创建工作簿对象
        sheet = workbook.add_sheet("sheet1")  # 创建工作表对象

        thStyle = xlwt.XFStyle()  # 创建表头样式对象
        thFont = xlwt.Font()  # 创建表头字体对象
        thFont.name = '楷体'
        thFont.bold = True
        thStyle.font = thFont  # 将设置表头字体样式添加到表样式的font变量中
        thAlign = xlwt.Alignment()  # 创建表头对齐方式对象
        thAlign.horz = 0x02  # 水平居中
        thAlign.vert = 0x01  # 垂直居中
        thStyle.alignment = thAlign
        sheet.write_merge(0, 1, 0, 1, "电影名称", thStyle)
        sheet.write_merge(0, 1, 2, 3, "电影评分", thStyle)
        sheet.write_merge(0, 1, 4, 5, "电影评论总数", thStyle)

        dataStyle = xlwt.XFStyle()  # 创建数据样式对象
        dataAlign = xlwt.Alignment()
        dataAlign.horz = 0x02
        dataAlign.vert = 0x01
        dataStyle.alignment = dataAlign
        dataFont = xlwt.Font()  # 创建数据字体对象
        dataFont.name = '楷体'
        dataStyle.font = dataFont

        n = len(self.filmScore)
        print(n)
        i = 2
        for item in self.filmScore.items():
            name = item[0]
            score = self.filmScore[name]
            number = self.numberOfReviews[name]
            sheet.write_merge(i, i, 0, 1, name, dataStyle)  # 合并单元格写入电影名称
            sheet.write_merge(i, i, 2, 3, score, dataStyle)  # 合并单元格写入评分
            sheet.write_merge(i, i, 4, 5, number, dataStyle)  # 合并单元格写入评论数
            i += 1
        workbook.save(path)

    '''
        将数据写入Excel中
    '''

    def writeExcel(self, sheet_name, save_path='.\\data\\filmData.xls'):
        workbook = xlwt.Workbook(encoding='ascii')  # 创建工作簿对象
        sheet = workbook.add_sheet(sheet_name)  # 创建工作表对象

        thStyle = xlwt.XFStyle()  # 创建表头样式对象
        thFont = xlwt.Font()  # 创建表头字体对象
        thFont.name = '楷体'
        thFont.bold = True
        thStyle.font = thFont  # 将设置表头字体样式添加到表样式的font变量中
        thAlign = xlwt.Alignment()  # 创建表头对齐方式对象
        thAlign.horz = 0x02  # 水平居中
        thAlign.vert = 0x01  # 垂直居中
        thStyle.alignment = thAlign
        sheet.write_merge(0, 1, 0, 1, "电影名称", thStyle)
        sheet.write_merge(0, 1, 2, 3, "电影评分", thStyle)
        sheet.write_merge(0, 1, 4, 5, "电影评论总数", thStyle)

        dataStyle = xlwt.XFStyle()  # 创建数据样式对象
        dataAlign = xlwt.Alignment()
        dataAlign.horz = 0x02
        dataAlign.vert = 0x01
        dataStyle.alignment = dataAlign
        dataFont = xlwt.Font()  # 创建数据字体对象
        dataFont.name = '楷体'
        dataStyle.font = dataFont

        n = len(self.filmScore)
        i = 2
        for item in self.filmScore.items():
            name = item[0]
            score = self.filmScore[name]
            number = self.numberOfReviews[name]
            sheet.write_merge(i, i, 0, 1, name, dataStyle)  # 合并单元格写入电影名称
            sheet.write_merge(i, i, 2, 3, score, dataStyle)  # 合并单元格写入评分
            sheet.write_merge(i, i, 4, 5, number, dataStyle)  # 合并单元格写入评论数
            i += 1
        workbook.save(save_path)
        time.sleep(0.1)
        print("数据写入成功,数据保存至" + save_path + "中")

    '''
        将数据追加到Excel中
    '''

    def writeExcelAppend(self, path):
        workbook = xlrd.open_workbook(path, formatting_info=True)  # formatting_info = True 保留原有格式
        sheets_name = workbook.sheet_names()
        sheet = workbook.sheet_by_name(sheets_name[0])
        rows = sheet.nrows
        new_workbook = copy(workbook)  # xlrd对象转为xlwt对象
        new_sheet = new_workbook.get_sheet(0)
        dataStyle = xlwt.XFStyle()  # 创建数据样式对象
        dataAlign = xlwt.Alignment()
        dataAlign.horz = 0x02
        dataAlign.vert = 0x01
        dataStyle.alignment = dataAlign
        dataFont = xlwt.Font()  # 创建数据字体对象
        dataFont.name = '楷体'
        dataStyle.font = dataFont
        i = rows
        for item in self.filmScore.items():
            name = item[0]
            score = item[1]
            number = self.numberOfReviews.get(name)
            new_sheet.write_merge(i, i, 0, 1, name, dataStyle)  # 合并单元格写入电影名称
            new_sheet.write_merge(i, i, 2, 3, score, dataStyle)  # 合并单元格写入评分
            new_sheet.write_merge(i, i, 4, 5, number, dataStyle)  # 合并单元格写入评论数
            i += 1
        new_workbook.save(path)
        time.sleep(0.1)
        print("数据追加成功,数据追加至" + str(path).split('\\')[2] + "文件中")


'''
    提取Excel表格数据
'''


class Utils:
    names = []
    comments = []

    def __init__(self, excel_path):
        workbook = xlrd.open_workbook(excel_path)
        self.sheet = workbook.sheets()[0]
        self.rows = int(self.sheet.nrows)

    def chooseData(self):
        # 随机从excel表中选出4个名称
        for i in range(1, 4 + 1):
            flag = random.randint(2, self.rows)
            self.names.append(self.sheet.cell_value(flag, 0))
            self.comments.append(self.sheet.cell_value(flag, 4))
        return self.names, self.comments

    @staticmethod
    def choosePage(page_number=1):
        print("爬取第" + str(page_number) + "页")
        start = 0  # 实际起始位置
        if page_number == 1:  # 第一页
            start = 0
            return start
        elif page_number == 2:  # 第二页
            start = 25
            return start
        elif page_number == 3:  # 第三页
            start = 50
            return start
        elif page_number == 4:  # 第四页
            start = 75
            return start
        elif page_number == 5:  # 第五页
            start = 100
            return start
        elif page_number == 6:  # 第六页
            start = 125
            return start
        elif page_number == 7:  # 第七页
            start = 150
            return start
        elif page_number == 8:  # 第八页
            start = 175
            return start
        elif page_number == 9:  # 第九页
            start = 200
            return start
        elif page_number == 10:  # 第十页
            start = 225
            return start
        else:
            raise '页码不合法'

    @staticmethod
    def write_log():
        log = open(".\\data\\logs.txt", 'a', encoding='utf-8')
        currentTime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
        log.write('\n' + str(currentTime) + "爬取数据！")
        log.close()


'''
    生成可视化文件
'''


class DataVisualization:
    def __init__(self, excel):
        self.path = excel
        workbook = xlrd.open_workbook(excel)
        self.sheet = workbook.sheets()[0]
        self.rows = self.sheet.nrows

    '''
        树形图
    '''

    def treeMap(self):
        workbook = xlrd.open_workbook(self.path)
        sheet = workbook.sheets()[0]
        rows = sheet.nrows
        flag = 0  # excel表有效数据标识符
        firstPage = []
        secondPage = []
        thirdPage = []
        fourthPage = []
        fifthPage = []
        sixthPage = []
        seventhPage = []
        eightPage = []
        ninthPage = []
        tenthPage = []
        for n in range(2, rows):
            if flag < 25:
                firstPage.append(
                    {"value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 25 < flag <= 50:
                secondPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 50 < flag <= 75:
                thirdPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 75 < flag <= 100:
                fourthPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 100 < flag <= 125:
                fifthPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 125 < flag <= 150:
                sixthPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 150 < flag <= 175:
                seventhPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 175 < flag <= 200:
                eightPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 200 < flag <= 225:
                ninthPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})
            elif 225 < flag <= 250:
                tenthPage.append(
                    {"score": sheet.cell_value(n, 2), "value": sheet.cell_value(n, 4), "name": sheet.cell_value(n, 0)})

            flag += 1
        data = [
            {"children": firstPage},
            {"children": secondPage},
            {"children": thirdPage},
            {"children": fourthPage},
            {"children": fifthPage},
            {"children": sixthPage},
            {"children": seventhPage},
            {"children": eightPage},
            {"children": ninthPage},
            {"children": tenthPage}
        ]
        c = (
            TreeMap()
                .add(series_name='', data=data)
                .set_global_opts(title_opts=opts.TitleOpts(title="豆瓣电影Top250"))
                .render(".\\data\\treeMap.html")
        )
        return 'TreeMap树状图生成成功'

    '''
        文字云
    '''

    def wordCloud(self):
        workbook = xlrd.open_workbook(self.path)
        sheet = workbook.sheets()[0]
        nrows = sheet.nrows
        data = []
        for n in range(2, nrows):
            data.append((sheet.cell_value(n, 0), sheet.cell_value(n, 4)))
        (
            WordCloud()
                .add(
                series_name="热点分析", data_pair=data,
                word_size_range=[6, 88], textstyle_opts=opts.TextStyleOpts(font_family="楷体"),
                shape=SymbolType.RECT)  # shape设置云形状
                .set_global_opts(
                title_opts=opts.TitleOpts(
                    title="热门电影文字云", title_textstyle_opts=opts.TextStyleOpts(font_size=23)
                )
            )
                .render(".\\data\\wordCloud.html")
        )
        return '文字云生成成功'

    '''
        散点图1
    '''

    def Scatter1(self):
        data = []
        for n in range(2, self.rows):
            data.append([self.sheet.cell_value(n, 2), self.sheet.cell_value(n, 4)])
        data.sort(key=lambda x: x[0])
        x_data = [d[0] for d in data]
        y_data = [d[1] for d in data]

        (
            Scatter(init_opts=opts.InitOpts(width="1600px", height="1000px"))
                .add_xaxis(xaxis_data=x_data)
                .add_yaxis(
                series_name="根据评分分析讨论最热数据",
                y_axis=y_data,
                symbol_size=20,
                label_opts=opts.LabelOpts(is_show=False),
            )
                .set_series_opts()
                .set_global_opts(
                xaxis_opts=opts.AxisOpts(
                    type_="value", splitline_opts=opts.SplitLineOpts(is_show=True)
                ),
                yaxis_opts=opts.AxisOpts(
                    type_="value",
                    axistick_opts=opts.AxisTickOpts(is_show=True),
                    splitline_opts=opts.SplitLineOpts(is_show=True),
                ),
                tooltip_opts=opts.TooltipOpts(is_show=False),
            )
                .render(".\\data\\scatter1.html")
        )
        return '散点图1生成成功'

    '''
        散点图2
    '''

    def Scatter2(self):
        u = Utils(excel_path='.\\data\\filmData.xls')
        data = u.chooseData()
        c = (
            Scatter()
                .add_xaxis(u.names)
                .add_yaxis("", u.comments)
                .set_global_opts(
                title_opts=opts.TitleOpts(title="电影评论数"),
                visualmap_opts=opts.VisualMapOpts(max_=3000000),
            )
                .render(".\\data\\scatter2.html")
        )
        return '散点图2生成成功'
