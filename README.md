# 抖店/飞鸽自动发送消息脚本

抖店 / 飞鸽 根据订单号自动发消息给客户，这两个脚本大大小小发送过好几次，视觉的至少发送过 5 次，后续因为朋友需要更快的速度才改用的 DrissionPage，理论上后者更优

当前项目有两个方案：

- 默认文件夹，使用 `DrissionPage` 进行自动化发送，优点就是速度快
- `vision`文件夹，使用 `pyautogui` 进行截图视觉化 + Umi OCR 进行发送，优点就是相对稳定

## 视觉发送说明

1. 下载比较稳定的 [Umi-OCR](https://github.com/hiroi-sora/Umi-OCR/releases)
2. 开启 `1224` 的 HTTP 端口
3. 运行程序

```shell
cd vision
python main.py
```

## 基于 DrissionPage 发送说明

1. 需要开启 `9111` debug 端口

2. 如果不是 `Edge 浏览器` 还需要自行修改一下 `browser_path` 和 `user_data_path`

```python
co = ChromiumOptions().set_paths(
    browser_path=r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',  # 请改为你电脑内Chrome可执行文件路径
    local_port=9111,
    user_data_path=r'C:\Users\Administrator\AppData\Local\Microsoft\Edge\User Data'
)
```

3. 运行程序

```shell
python main.py
```