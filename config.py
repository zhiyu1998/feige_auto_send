import os
from enum import Enum

# 使用相对路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
excel_data_path = os.path.join(BASE_DIR, "test.xlsx")
output_excel_path = os.path.join(BASE_DIR, "output.xlsx")
send_messages = {
    "video": os.path.join(BASE_DIR, "demo.mp4"),
    "image": os.path.join(BASE_DIR, "demo.png"),
    "text": "亲亲，上面是垂丝茉莉养护方式，请您查收"
}


class Status(Enum):
    ONLINE = 0
    BREAK = 1


switch_status = Status.BREAK
