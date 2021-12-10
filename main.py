import request
import requests
req = request.Request()
result = req.get(url="https://movie.douban.com/top250")
if result:
    req.do()
else:
    print("出错！")

