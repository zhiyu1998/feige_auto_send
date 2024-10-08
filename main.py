import os
import time
import numpy as np

from DrissionPage import Chromium, ChromiumOptions
from DrissionPage.common import Keys

from config import send_messages, excel_data_path, switch_status, output_excel_path
from logger_config import logger
from utils import read_excel, load_processed_clients, save_to_csv, read_csv

# 读取数据
excel_order_nums = [str(*excel_order_num).replace('\t', '') for excel_order_num in read_excel(excel_data_path).values]

logger.info(f'共计有{len(excel_order_nums)}条待发送消息')

# 加载已处理的客户
client_set = load_processed_clients(output_excel_path)
client_data = read_csv(output_excel_path).values
# logger.info(client_data)
logger.info(f'加载了{len(client_set)}个已处理过的客户')

# 这里做数据差集，去除已处理的客户
excel_order_nums = list(set(excel_order_nums) - client_set)
logger.info(f'将从{excel_order_nums[0]}开始发送')
logger.info(f'差集后剩余{len(excel_order_nums)}条待发送消息')
# logger.info("6917920175080767044" in excel_order_nums)
# input("Press any key to start")

path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'  # 请改为你电脑内Chrome可执行文件路径
co = ChromiumOptions().set_paths(
    local_port=9111,
    user_data_path=r'C:\Users\Administrator\AppData\Local\Microsoft\Edge\User Data'
)

tab = Chromium(addr_or_opts=co).latest_tab
# 打开会话
tab.get('https://im.jinritemai.com/pc_seller_v2/main/workspace')

# 判断是否是离线
offline_sign = tab.ele('xpath=//*[@id="rootContainer"]/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div')
tab.wait.ele_displayed(offline_sign)
time.sleep(2)
logger.info(f'当前状态为：<{offline_sign.text}>')
if "离线" == offline_sign.text:
    offline_sign.click()
    # 切换到在线
    switch_online = None
    if switch_status == 0:
        switch_online = tab.ele('xpath=//*[@id="rootContainer"]/div/div[1]/div[1]/div[1]/div[1]/div/div[2]/div[3]')
    elif switch_status == 1:
        switch_online = tab.ele('xpath=//*[@id="rootContainer"]/div/div[1]/div[1]/div[1]/div[1]/div/div[2]/div[4]')
    tab.wait.ele_displayed(switch_online)
    switch_online.click()
    logger.info(f"已切换至 <{switch_status == 0 and '在线' or '小休'} 状态>，开始上班")

for excel_order_num in excel_order_nums:
    # 点击订单号
    order_num = tab.ele('xpath=//*[@id="rootContainer"]/div/div[1]/div[1]/div[2]/div/input')
    tab.wait.ele_displayed(order_num)
    order_num.click()
    order_num.input(excel_order_num)

    # 获取用户名称并点击
    user_info = tab.ele('xpath=//*[@id="chantListScrollArea"]/div[2]')
    while "来自订单" not in user_info.text:
        user_info = tab.ele('xpath=//*[@id="chantListScrollArea"]/div[2]')
        time.sleep(0.5)
    tab.wait.ele_displayed(user_info)
    logger.info(f'正在处理订单号: {user_info.text}')
    user_info.click()

    # 获取客户昵称
    time.sleep(2)
    client_name = tab.ele('xpath=//*[@id="workspace-chat"]/div[3]/div[1]/div[1]/span/div').text

    logger.info(f'当前客户：{client_name}，正在加载聊天记录：')
    # 获取聊天记录
    chat_histroy = tab.ele('xpath=//*[@id="workspace-chat"]/div[3]/div[3]/div/div[2]').children()

    cur_history = []
    # 发送过标志
    is_sended = False
    # 通过遍历聊天记录判断是否发送过
    if len(chat_histroy) != 1:
        for chat_one in chat_histroy:
            # //*[@id="workspace-chat"]/div[3]/div[3]/div/div[2]/div[9]/div/div/div[2]/div[2]/div[1]/pre/span
            cur_history.append(chat_one.text)
        for chat_one in cur_history:
            if send_messages.get("text") in chat_one:
                is_sended = True
                break

    logger.info(f'是否已经发送过：{is_sended}')

    status = "跳过"
    # 如果已经发送过text就跳过
    if is_sended:
        logger.info(f'客户：{client_name} 已处理过，跳过')
    else:
        # set不存在客户，添加进set
        client_set.add(client_name)
        status = "发送"

        try:
            for message_type, message in send_messages.items():
                if message_type == 'text':
                    logger.info(f'正在发送消息：{message}')
                    # 点击输入框
                    feige_input = tab.ele('xpath=//*[@id="workspace-chat"]/div[3]/div[6]/textarea')
                    # tab.wait.ele_displayed(feige_input)
                    feige_input.click()
                    feige_input.input(message)
                    time.sleep(1)
                    # 点击发送
                    feige_send = tab.ele('xpath=//*[@id="workspace-chat"]/div[3]/div[6]/div[3]/div[3]/div')
                    feige_send.click()

                elif message_type == 'image':
                    logger.info(f'正在发送图片：{message}')
                    # 点击图片图标
                    feige_image = tab.ele('xpath=//*[@id="workspace-chat"]/div[3]/div[6]/div[2]/div[1]/label[1]')
                    feige_image.click.to_upload(os.path.abspath(message))
                    logger.info(f"开始上传图片{os.path.abspath(message)}")
                    # 点击发送
                    time.sleep(1)
                    feige_send = tab.ele('xpath=/html/body/div[6]/div/div[3]/div[3]')
                    tab.wait.ele_displayed(feige_send)
                    feige_send.click()

                elif message_type == 'video':
                    feige_video = tab.ele('xpath=//*[@id="workspace-chat"]/div[3]/div[6]/div[2]/div[1]/label[2]')
                    feige_video.click.to_upload(os.path.abspath(message))
                    logger.info(f"开始上传视频{os.path.abspath(message)}")
                    time.sleep(1)
                    feige_send = tab.ele('xpath=/html/body/div[6]/div/div[2]/div/div[2]/div[3]/button[1]')
                    tab.wait.ele_displayed(feige_send)
                    feige_send.click()

                time.sleep(0.5)

        except Exception as e:
            logger.info(f"发送消息给 {client_name} 时出错：{str(e)}")
            status = "失败"

    time.sleep(0.5)
    # 关闭客户
    tab.actions.key_down(Keys.ESCAPE)
    tab.actions.key_up(Keys.ESCAPE)
    # 添加客户数据到列表
    np.vstack((client_data, np.array([excel_order_num, client_name, status])))

    # 每处理完一个客户后保存数据
    save_to_csv(output_excel_path, client_data)

    logger.info("-----------------------------")
    time.sleep(0.5)

logger.info("所有订单处理完毕")
