import os
from enum import Enum

# 使用相对路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Excel 路径
excel_data_path = os.path.join(BASE_DIR, "test.xlsx")
# 输出 Excel 路径
output_excel_path = os.path.join(BASE_DIR, "output.csv")
# 发送消息，按照顺序发送（video-视频、image-图片、text-文字）
send_messages = {
    "video": os.path.join(BASE_DIR, "demo.mp4"),
    "image": os.path.join(BASE_DIR, "demo.png"),
    "text": "亲亲，由于采购过来的绣球颜色较少，所以可能给您发的随机绣球颜色会相同~但都是高品质的好绣球，这边先发给您养护方式，您收到有任何问题都可以联系我们，再次抱歉给您添麻烦了"
}


# 登录时切换的状态（0-在线，1-小休）
class Status(Enum):
    ONLINE = 0
    BREAK = 1


# 登录时切换的状态（ONLINE-在线，BREAK-小休）
switch_status = Status.BREAK

# 是否使用无序集合来提高效率
use_unordered_set = False
