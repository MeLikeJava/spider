import request
import time
req = request.Request()
status = req.get(url="https://movie.douban.com/top250", start=25)
save_path = ".\\data\\filmData.xls"
if status:
    req.do(path=save_path)
else:
    print("出错！")

