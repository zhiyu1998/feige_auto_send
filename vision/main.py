import time
import logging
from utils import AutomationTool
from config import *


def process_order(order_num):
    """
    处理订单
    @param order_num: 订单号
    """
    # 找到订单
    AutomationTool.screenshot_and_click(TARGET1)
    AutomationTool.type_text(order_num)  # 输入文字

    # 点击来自订单的那个人（需要time.sleep等待一下）
    time.sleep(1)
    is_find_order = AutomationTool.screenshot_and_click(f"来自订单：{str(order_num)[:8]}")

    # 无效订单号, 找不到就下一个
    if not is_find_order:
        # 点击关闭重置订单
        AutomationTool.click_image(IMAGE_SHUTDOWN_BTN_PATH)
        return False

    # 再次截图，这个截图要同时判断是否已经发送了订单购买人 && 是否重复订单人
    if check_message_sent():
        AutomationTool.click_image(IMAGE_SHUTDOWN_BTN_PATH)
        return False

    send_message()
    return True


def check_message_sent():
    """
    检查是否已经发送了订单购买人的信息
    """
    ocr_data = AutomationTool.process_screenshot_for_ocr()
    if ocr_data:
        # 遍历所有识别到的文字，判断是否已经包含了发送的消息
        for item in ocr_data['data']:
            if SENT_MESSAGE_TEXT in item['text']:
                logging.info("已经发送过订单购买人的信息，跳过本次循环")
                return True
    return False


def send_message():
    """
    发送消息
    """
    # 这里防止出现问题，点击下商店
    # AutomationTool.click_image(IMAGE_STORE_BTN_PATH)
    time.sleep(0.5)

    # 查找并点击图像
    if IMAGE_PATH.endswith(("png", "jpg")):
        AutomationTool.click_image(IMAGE_BTN_PATH)
    else:
        AutomationTool.click_image(VIDEO_BTN_PATH)
    AutomationTool.type_text(IMAGE_PATH)
    AutomationTool.press_enter()
    # 如果是视频就点击发送
    if IMAGE_PATH.endswith("mp4"):
        AutomationTool.click_image(SEND_VIDEO_BTN_PATH)
    time.sleep(0.5)
    AutomationTool.press_enter()

    AutomationTool.screenshot_and_click(TARGET2)
    AutomationTool.type_text(SENT_MESSAGE_TEXT)
    AutomationTool.press_enter()

    # 关闭订单
    AutomationTool.press_esc()


def main():
    order_nums = AutomationTool.read_excel(EXCEL_PATH).values
    total_orders = len(order_nums)
    processed_orders = 0
    successful_orders = 0

    try:
        for order_num in order_nums:
            time.sleep(0.5)
            # 获取订单号
            processed_orders += 1
            logging.info(f"处理订单 {processed_orders}/{total_orders}: {order_num[0]}")
            if process_order(order_num[0]):
                successful_orders += 1
    except Exception as e:
        logging.error(f"程序执行出错：{str(e)}")
    finally:
        logging.info(f"程序执行完毕。成功处理 {successful_orders}/{total_orders} 个订单。")


if __name__ == '__main__':
    # ！！！ 输入这个激活环境：.\venv\Scripts\activate.ps1
    main()
