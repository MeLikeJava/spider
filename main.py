import request
import tools as t
req = request.Request()  # 实例化一个request模块中的Request类对象

page = t.Utils.choosePage(page_number=2)  # page_number:参数取值范围1~10, 默认为第一页

# 通过Request对象req调用其中自定义的get()函数，目的是向豆瓣网服务器发送请求，会产生一个boolean类型的返回值，用于下面判断使用
status = req.get(url="https://movie.douban.com/top250", page=page)

# 设定数据保存路径,就是将网页爬取的数据保存至那个地方的放个excel数据表格中
save_path = ".\\data\\filmData.xls"
if status:
    req.do(path=save_path)  # 执行Request类中的do()函数
else:
    print("出错！")

