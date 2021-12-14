import request
import tools as t
req = request.Request()

page = t.Utils.choosePage(page_number=1)  # page_number:参数取值范围1~10, 默认为第一页

status = req.get(url="https://movie.douban.com/top250", page=page)
save_path = ".\\data\\filmData.xls"
if status:
    req.do(path=save_path)
else:
    print("出错！")

