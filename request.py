import requests
import re
import tools as t


class Request:
    # 类成员html
    html = ''

    def get(self, url):
        # 模拟http头文件信息
        head = {
            "USER-Agent": 'Mozilla / 5.0(Windows NT 10.0; Win64; x64) AppleWebKit / 537.36(KHTML, like Gecko) Chrome '
                          '/ 80.0.3987.122 Safari / 537.36 '
        }
        html = requests.get(url, timeout=30, headers=head)
        html.encoding = 'utf-8'
        if html.status_code != 200:
            print("请求失败,请检查！")
            return False
        self.html = html.text
        return True

    def do(self):
        obj = t.DataResolve(html=self.html)  # 实例化DataResolve对象
        result = obj.resolve()  # 执行解析
        print(result)
