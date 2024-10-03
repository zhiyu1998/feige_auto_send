import os
import time

from DrissionPage import Chromium, ChromiumOptions
from DrissionPage.common import Keys
from DrissionPage.errors import ElementNotFoundError

from config import send_messages, excel_data_path, switch_status, output_excel_path, Status
from logger_config import logger
from utils import read_excel, save_to_excel, load_processed_clients

# 读取数据
excel_order_nums = read_excel(excel_data_path).values

# 加载已处理的客户
client_set = load_processed_clients(output_excel_path)
client_data = []

path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'  # 请改为你电脑内Chrome可执行文件路径
co = ChromiumOptions().set_paths(
    local_port=9111,
    user_data_path=r'C:\Users\Administrator\AppData\Local\Microsoft\Edge\User Data'
)

with Chromium(addr_or_opts=co) as browser:
    tab = browser.latest_tab
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
        if switch_status == Status.ONLINE:
            switch_online = tab.ele('xpath=//*[@id="rootContainer"]/div/div[1]/div[1]/div[1]/div[1]/div/div[2]/div[3]')
        elif switch_status == Status.BREAK:
            switch_online = tab.ele('xpath=//*[@id="rootContainer"]/div/div[1]/div[1]/div[1]/div[1]/div/div[2]/div[4]')
        tab.wait.ele_displayed(switch_online)
        switch_online.click()
        logger.info(f"已切换至 <{switch_status == Status.ONLINE and '在线' or '小休'} 状态>，开始上班")


    def wait_and_click(tab, xpath):
        element = tab.ele(f'xpath={xpath}')
        tab.wait.ele_displayed(element)
        element.click()
        return element


    def send_message(tab, message_type, message):
        if message_type == 'text':
            feige_input = wait_and_click(tab, '//*[@id="workspace-chat"]/div[3]/div[6]/textarea')
            feige_input.input(message)
            wait_and_click(tab, '//*[@id="workspace-chat"]/div[3]/div[6]/div[3]/div[3]/div')
        elif message_type == 'image':
            feige_image = wait_and_click(tab, '//*[@id="workspace-chat"]/div[3]/div[6]/div[2]/div[1]/label[1]')
            feige_image.click.to_upload(os.path.abspath(message))
            wait_and_click(tab, '/html/body/div[6]/div/div[3]/div[3]')
        elif message_type == 'video':
            feige_video = wait_and_click(tab, '//*[@id="workspace-chat"]/div[3]/div[6]/div[2]/div[1]/label[2]')
            feige_video.click.to_upload(os.path.abspath(message))
            wait_and_click(tab, '/html/body/div[6]/div/div[2]/div/div[2]/div[3]/button[1]')


    def wait_for_element_with_text(tab, xpath, text, timeout=10):
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                element = tab.ele(f'xpath={xpath}')
                if text in element.text:
                    return element
            except ElementNotFoundError:
                pass
            time.sleep(0.5)
        raise TimeoutError(f"标签 '{text}' 在 {timeout} 秒内没有找到")


    for excel_order_num in excel_order_nums:
        # 解构
        [excel_order_num] = excel_order_num

        # 点击订单号
        order_num = tab.ele('xpath=//*[@id="rootContainer"]/div/div[1]/div[1]/div[2]/div/input')
        tab.wait.ele_displayed(order_num)
        order_num.click()
        order_num.input(excel_order_num)

        # 获取用户名称并点击
        try:
            user_info = wait_for_element_with_text(tab, '//*[@id="chantListScrollArea"]/div[2]', "来自订单")
            logger.info(f'正在处理订单号: {user_info.text}')
            user_info.click()
        except TimeoutError:
            logger.warning(f"未能找到包含'来自订单{excel_order_num}'的用户信息")
            continue

        # 获取客户昵称
        client_name = tab.ele('xpath=//*[@id="workspace-chat"]/div[3]/div[1]/div[1]/span/div').text
        logger.info(f'当前客户：{client_name}')

        status = "跳过"
        if client_name in client_set:
            logger.info(f'客户：{client_name} 已处理过，跳过')
        else:
            # set不存在客户，添加进set
            client_set.add(client_name)
            status = "发送"

            try:
                for message_type, message in send_messages.items():
                    logger.info(
                        f'正在发送{"消息" if message_type == "text" else "图片" if message_type == "image" else "视频"}：{message}')
                    send_message(tab, message_type, message)
                    time.sleep(0.5)

            except Exception as e:
                logger.error(f"发送消息给 {client_name} 时出错：{str(e)}")
                status = "失败"
                # 失败后保存数据
                client_data.append([excel_order_num, client_name, status])
                save_to_excel(output_excel_path, client_data)

        time.sleep(0.5)
        tab.actions.key_down(Keys.ESCAPE)
        tab.actions.key_up(Keys.ESCAPE)
        # 添加客户数据到列表
        client_data.append([excel_order_num, client_name, status])

        # 每处理完一个客户后保存数据
        save_to_excel(output_excel_path, client_data)

        logger.info("-----------------------------")
        time.sleep(0.5)

logger.info("所有订单处理完毕")
