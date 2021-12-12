import requests
import re
import tools as t
import xlrd
import time

class Request:
    # 类成员html
    html = ''

    def get(self, url, start):
        # 模拟http头文件信息
        head = {
            "USER-Agent": 'Mozilla / 5.0(Windows NT 10.0; Win64; x64) AppleWebKit / 537.36(KHTML, like Gecko) Chrome '
                          '/ 80.0.3987.122 Safari / 537.36 '
        }
        url = url + '?start=' + str(start)  # 组装url地址
        html = requests.get(url, timeout=30, headers=head)
        html.encoding = 'utf-8'
        if html.status_code != 200:
            print("请求失败,请检查！")
            return False
        self.html = html.text
        return True

    def do(self, path):
        obj = t.DataResolve(html=self.html)  # 实例化DataResolve对象
        print("*"*20, "解析执行开始", "*"*20)
        result = obj.resolve()  # 执行解析
        time.sleep(0.1)
        print("*"*20, "解析结束", "*"*20)
        workbook = xlrd.open_workbook(path)
        sheets_name = workbook.sheet_names()
        sheet = workbook.sheet_by_name(sheet_name=sheets_name[0])
        rows = sheet.nrows
        sheet_name = "工作表1"
        print("*" * 20, "数据写入执行开始", "*" * 20)
        if rows == 0:
            obj.writeExcel(sheet_name=sheet_name, save_path=path)
            time.sleep(0.1)
            print("*" * 20, "数据写入结束", "*" * 20)
            dv = t.DataVisualization(excel=".\\data\\filmData.xls")
            print(dv.wordCloud())
        else:
            obj.writeExcelAppend(path=path)
            time.sleep(0.1)
            print("*" * 20, "数据写入结束", "*" * 20)
            dv = t.DataVisualization(excel=".\\data\\filmData.xls")
            print(dv.wordCloud())

