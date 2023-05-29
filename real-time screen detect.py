import torch
import numpy as np
import cv2
from PIL import ImageGrab
from pathlib import Path
import time
import pyautogui
import math
import win32gui
import win32api
import win32con
import time
import wave
import pyaudio



def get_direction(center):    #用来获取当前检测目标在鼠标的哪个方位的函数
    current = pyautogui.position()
    if center[0]<current[0] and center[1]>current[1]:
        str0 = 'zuoshang'
    elif center[0]>current[0] and center[1]>current[1]:
        str0 = 'youshang'
    elif center[0]<current[0] and center[1]<current[1]:
        str0 = 'zuoxia'
    elif center[0]>current[0] and center[1]<current[1]:
        str0 = 'youxia'
    return str0


def play(filename):    #启动语音播放函数
    chunk = 1024
    wf = wave.open(filename, 'rb')
    p = pyaudio.PyAudio()
    stream = p.open(format=p.get_format_from_width(wf.getsampwidth()), channels=wf.getnchannels(),
                    rate=wf.getframerate(), output=True)

    data = wf.readframes(chunk)  # 读取数据
    # print(data)
    while data != b'':  # 播放
        stream.write(data)
        data = wf.readframes(chunk)
        # print('while循环中！')
        # print(data)
    stream.stop_stream()  # 停止数据流
    stream.close()
    p.terminate()  # 关闭 PyAudio


# print("Window %s:" % win32gui.GetWindowText(hwnd))
# print("\tLocation: (%d, %d)" % (x, y))
# print("\t    Size: (%d, %d)" % (w, h))
# TENSORRT_ENGINE_PATH_PY = 'Mine.engine'
# logger = trt.Logger(trt.Logger.WARNING)
# with open(TENSORRT_ENGINE_PATH_PY, "rb") as f, trt.Runtime(logger) as runtime:
#     engine = runtime.deserialize_cuda_engine(f.read())


def get_closet_center(detections,choice):    #获取距离鼠标最近的那个目标的中心位置的函数
    leftmost_center = None
    leftmost_center1 = None
    leftmost_x = float('inf')
    most = 100000
    current_position = pyautogui.position()
    if choice:    #如果choice不是空的，那么启动追踪模式后只会追踪之前输入的几个参数对应的目标
        for det in detections:
            cls = det[-1]
            if cls in choice:    #找到目前所找到的所有目标中离鼠标最近的那个目标的中心位置坐标
                x1, y1, x2, y2 = det[:4]
                center_x = (x1 + x2) / 2
                center_y = (y1 + y2) / 2
                leftmost_center = (center_x.item(), center_y.item())
                delta_x = (leftmost_center[0] - current_position[0])
                delta_y = (leftmost_center[1] - current_position[1])
                add = (delta_x**2 + delta_y**2)
                distances = math.sqrt(add)
                if distances < most:
                    most = distances
                    leftmost_center1 = leftmost_center
    else:    #如果choice是空列表那么会追踪所有检测到的东西
        for det in detections:
            cls = det[-1]
            if cls in [0,1,2,3,4,5,6,7]:
                x1, y1, x2, y2 = det[:4]
                center_x = (x1 + x2) / 2
                center_y = (y1 + y2) / 2
                leftmost_center = (center_x.item(), center_y.item())
                delta_x = (leftmost_center[0] - current_position[0])
                delta_y = (leftmost_center[1] - current_position[1])
                add = (delta_x**2 + delta_y**2)
                distances = math.sqrt(add)
                if distances < most:
                    most = distances
                    leftmost_center1 = leftmost_center
    return leftmost_center1    #返回值是离鼠标最近那个目标的中心位置所对应的坐标
def get_closet_center2(detections,choice):    #获取距离鼠标最近的那个目标的中心位置的函数
    leftmost_center = None
    leftmost_center1 = None
    leftmost_x = float('inf')
    least = 0
    if choice:    #如果choice不是空的，那么启动追踪模式后只会追踪之前输入的几个参数对应的目标
        for det in detections:
            cls = det[-1]
            if cls in choice:    #找到目前所找到的所有目标中离鼠标最近的那个目标的中心位置坐标
                x1, y1, x2, y2 = det[:4]
                center_x = (x1 + x2) / 2
                center_y = (y1 + y2) / 2
                leftmost_center = (center_x.item(), center_y.item())
                if center_x > least:
                    least = center_x
                    leftmost_center1 = leftmost_center
    else:    #如果choice是空列表那么会追踪所有检测到的东西
        for det in detections:
            cls = det[-1]
            if cls in [0,1,2,3,4,5,6,7]:
                x1, y1, x2, y2 = det[:4]
                center_x = (x1 + x2) / 2
                center_y = (y1 + y2) / 2
                leftmost_center = (center_x.item(), center_y.item())
                if center_x > least:
                    least = center_x
                    leftmost_center1 = leftmost_center
    return leftmost_center1    #返回值是离鼠标最近那个目标的中心位置所对应的坐标
def updown():
    pyautogui.move(0, 180, duration=0.1)
    time.sleep(0.6)
    win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
    pyautogui.move(0, -180, duration=0.1)

#调用模型和设置参数
FILE = Path(__file__).resolve()       #YOLOv5母目录的路径
ROOT = str(FILE.parents[0])   #根路径
windowclass = 'LWJGL'    #窗口类型
windowname = 'Minecraft 神秘6的复兴'    #窗口名称

choice1 = []
choice2 = []
choice3 = []#如果啥矿都想挖就不填，不然就填数字：0煤炭1铁2铜3红石4青金石5钻石6银7金
cho1 = 1
cho2 = 1
while cho1 != 0:#接受来自键盘，用来指明希望检测的矿
    cho1 = int(input("请输入你想要挖的矿的种类：1煤炭 2铁 3铜 4红石 5青金石 6钻石 7银 8金,如果不需要添加了就输入0: "))
    choice1.append(int(cho1 - 1))
while cho2 != 0:#接受来自键盘，用来指明希望检测的树种类
    cho2 = int(input("请输入你想要砍的树的种类：1橡木 2白桦 3合金欢 4云杉木,如果不需要添加了就输入0: "))
    choice2.append(int(cho2 - 1))

model1 = torch.hub.load(ROOT, "custom", path=ROOT + "\\kuang.pt", source="local")    #模型1：挖矿
model1.conf = 0.55  # confidence threshold (0-1)    #最大置信度，高于最大置信度则被检测为detections类
model1.iou = 0.45 #iou损失函数

model2 = torch.hub.load(ROOT, "custom", path=ROOT + "\\Wood.pt", source="local")    #树木检测模型
model2.conf = 0.8
model2.iou = 0.45

model3 = torch.hub.load(ROOT, "custom", path=ROOT + "\\monsters.pt", source="local")    #怪物检测模型
model3.conf = 0.75
model3.iou = 0.45
print("硕得岛离为您播报：模型加载完毕！")
play('loaded.wav')
cv2.namedWindow("detect", cv2.WINDOW_NORMAL)    #先定义一个名为detect的窗口，之后引入模型再用到此窗口，用来显示实时检测的画面
cv2.resizeWindow("detect", 800, 450)    #窗口的大小
previous_time = 0    #用来计算帧率
start = 2    #一个用来判断是否启用追踪模式的参数
fps = 2    #一个用来判断是否启用显示帧率的参数
mode = 0    #共有挖矿，砍树，砍怪，射箭四个模式
rotate = 2    #一个用来检测是否开始原地旋转检测矿物的参数
upordown = 2    #一个用来检测是否开始开启地底挖矿的参数
add_x = 0
add_x1 = 0
add_y = 0
turn = 0
d = 0
ticks = 0
ticks1 = 0
while True:    #循环执行
    hwnd = win32gui.FindWindow(windowclass, windowname)    #用win32gui函数来获取窗口
    rect = win32gui.GetWindowRect(hwnd)    #根据hwnd截取的窗口获取其大小
    x = rect[0]    #窗口左上角那个点在屏幕里的x坐标
    y = rect[1]    #窗口左上角那个点在屏幕里的y坐标
    w = rect[2] - x    #窗口的宽度
    h = rect[3] - y    #窗口的高度
    screen = np.array(ImageGrab.grab(bbox=(x, y, w, h)))    #利用imagegrab函数获取当前窗口的截屏，并转化为numpy矩阵（便于后面的RGB转化）
    if not cv2.getWindowProperty("detect", cv2.WND_PROP_VISIBLE):    #如果关掉detect窗口那么程序结束
        cv2.destroyAllWindows()
        exit("程序结束")
    if (win32api.GetAsyncKeyState(0x31) & 0x8000) > 0:    #按下主键盘1键启动挖矿模式
        print("mining mod start up!")
        # play('otto3.wav')
        mode = 1
        time.sleep(0.01)
    if (win32api.GetAsyncKeyState(0x32) & 0x8000) > 0:    #按下主键盘2键启动伐木模式
        print("felling mod start up!")
        mode = 2
        time.sleep(0.01)
    if (win32api.GetAsyncKeyState(0x33) & 0x8000) > 0:  # 按下主键盘3键启动打怪模式
        print("kill mod start up!")
        mode = 3
        time.sleep(0.01)
    if (win32api.GetAsyncKeyState(0x34) & 0x8000) > 0:  # 按下主键盘4启动射箭模式
        print("arrow mod start up!")
        mode = 4
        time.sleep(0.01)
    if (win32api.GetAsyncKeyState(0x35) & 0x8000) > 0:  # 按下主键盘5空档
        print("resetting")
        mode = 5
        time.sleep(0.01)
    if (win32api.GetAsyncKeyState(0x36) & 0x8000) > 0:  # 按下主键盘6开转
        rotate = rotate + 1
        time.sleep(0.01)
        print("rotate!")
    if (win32api.GetAsyncKeyState(0x37) & 0x8000) > 0:  # 按下主键盘7开地底挖矿
        upordown = upordown + 1
        time.sleep(0.01)
        print("underground!")
    if mode == 1:    #挖矿模式，按下鼠标中键开始追踪矿物（设了参数只有当鼠标离目标足够近才会开始自动挖矿），按下6开始转圈圈扫描模式，寻找周围是否有choice1里面的矿
        if (win32api.GetAsyncKeyState(0x04) & 0x8000) > 0:    #鼠标中键启动追踪模式
            start = start + 1    #奇数开偶数关，如果检测到按下中键，就奇数变偶数偶数变奇数
            time.sleep(0.5)
            if start % 2 == 0 and start != 2:    #奇数开偶数关
                print("link close o.0")
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                play('otto2.wav')
                win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
            if start % 2 != 0:    #奇数开偶数关
                print("link start!")
                play('otto.wav')
                win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
        choice = choice1    #挖矿模式要挖的东西
        results = model1(screen)    #根据模型来进行目标检测
        nearest_center = get_closet_center2(results.xyxy[0], choice)    #根据获取的detections类找到其中离当前鼠标位置最近的那个目标
        current_position = pyautogui.position()    #当前鼠标位置
        results = np.squeeze(results.render())    #用squeeze函数把矩阵还原为图片
        results = cv2.cvtColor(results, cv2.COLOR_BGR2RGB)    #把BGR还原成RGB图像
        cv2.imshow("detect", results)    #在原先的detect窗口输出实时的检测结果
        cv2.waitKey(1)

        if start % 2 != 0:    #追踪模式奇数启动偶数关闭
            if nearest_center:
                delta_x = (nearest_center[0] - current_position[0]) / 2.5
                delta_y = (nearest_center[1] - current_position[1]) / 2.5

                if upordown % 2 != 0:
                    limit_y = 88
                    limit_x = 88
                else:
                    limit_y = 10000
                    limit_x = 10000

                if abs(delta_y) < limit_y:
                    if abs(delta_x) < limit_x:
                        pyautogui.move(delta_x, delta_y, duration=0.1)
                        add_x = add_x + delta_x
                        add_y = add_y + delta_y
                        d = math.sqrt(delta_x**2 + delta_y**2)
                        if turn == 0 and d < 10:
                            turn = 1
                            win32api.keybd_event(87, 0, 0, 0)
                            time.sleep(0.05)
                            # ticks1 = ticks1 + 1
                            win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
                        if turn == 1 and d < 5:
                            if ticks % 15000 == 0:
                                turn = 0
                            else:
                                ticks = ticks + 1
                            win32api.keybd_event(9, 0, 0, 0)
                            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                            time.sleep(0.6)
                            win32api.keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
                            if upordown % 2 != 0:
                                updown()
                            d = 10000
                        elif turn == 1 and d > 5:
                            if ticks % 40000 == 0:
                                turn = 0
                            else:
                                ticks = ticks + 1
                            # win32api.keybd_event(83, 0, 0, 0)
                            # time.sleep((0.05))
                            # ticks1 = 0
                            # win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)
                            # # add_x = -add_x
                            # # add_y = -add_y
                            # # pyautogui.move(add_x, add_y, duration=0.01)
                            # turn = 0
                            # # add_x = 0
                            # # add_y = 0
                    else:
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                        time.sleep(0.02)
            elif rotate % 2 != 0:
                # win32api.keybd_event(87, 0, 0, 0)
                # time.sleep(0.1)
                # win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
                d = 10000
                ticks = 0
                if turn == 1:
                    add_x1 = 0
                    turn = 0
                pyautogui.move(-100, 0, duration=0.05)
                add_x1 = add_x1 - 100
                if add_x1 % (-2400) == 0:    #转一圈就暂停然后向前走一段距离
                    win32api.keybd_event(87, 0, 0, 0)
                    time.sleep(0.2)
                    win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

                time.sleep(0.02)
            else:
                add_x1 = 0
        else:
            add_x1 = 0
                # else:
                    # pyautogui.move(0, -116, duration=0.5)
                    # # time.sleep(0.8)
                    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                    # # time.sleep(1)
                    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                    # pyautogui.move(0, 116, duration=0.5)
                    # # time.sleep(0.1)
                    # win32api.keybd_event(57, 0, 0, 0)
                    # win32api.keybd_event(57, 0, win32con.KEYEVENTF_KEYUP, 0)
                    # # time.sleep(0.12)
                    # pyautogui.move(0, 116, duration=0.5)
                    # # time.sleep(0.1)
                    # win32api.keybd_event(32, 0, 0, 0)
                    # time.sleep(0.1)
                    # win32api.keybd_event(32, 0, win32con.KEYEVENTF_KEYUP, 0)
                    # # time.sleep(0.05)
                    # win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
                    # time.sleep(0.1)
                    # win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)
                    # # time.sleep(0.7)
                    # win32api.keybd_event(32, 0, win32con.KEYEVENTF_KEYUP, 0)
                    # pyautogui.move(0, -116, duration=0.5)
                    # # time.sleep(0.1)
                    # win32api.keybd_event(49, 0, 0, 0)
                    # # time.sleep(0.12)

                # else:
    if mode == 2:    #挖矿模式，按下鼠标中键开始追踪矿物（设了参数只有当鼠标离目标足够近才会开始自动挖矿），按下6开始转圈圈扫描模式，寻找周围是否有choice1里面的矿
        if (win32api.GetAsyncKeyState(0x04) & 0x8000) > 0:    #鼠标中键启动追踪模式
            start = start + 1    #奇数开偶数关，如果检测到按下中键，就奇数变偶数偶数变奇数
            time.sleep(0.5)
            if start % 2 == 0 and start != 2:    #奇数开偶数关
                print("link close o.0")
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                play('otto2.wav')
                win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
            if start % 2 != 0:    #奇数开偶数关
                print("link start!")
                play('otto.wav')
                win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
        choice = choice2    #挖矿模式要挖的东西
        results = model2(screen)    #根据模型来进行目标检测
        nearest_center = get_closet_center(results.xyxy[0], choice)    #根据获取的detections类找到其中离当前鼠标位置最近的那个目标
        current_position = pyautogui.position()    #当前鼠标位置
        results = np.squeeze(results.render())    #用squeeze函数把矩阵还原为图片
        results = cv2.cvtColor(results, cv2.COLOR_BGR2RGB)    #把BGR还原成RGB图像
        cv2.imshow("detect", results)    #在原先的detect窗口输出实时的检测结果
        cv2.waitKey(1)

        if start % 2 != 0:    #追踪模式奇数启动偶数关闭
            if nearest_center:
                delta_x = (nearest_center[0] - current_position[0]) / 2.5
                delta_y = (nearest_center[1] - current_position[1]) / 2.5
                add_x = add_x + delta_x
                add_y = add_y + delta_y
                if upordown % 2 != 0:
                    limit_y = 88
                    limit_x = 88
                else:
                    limit_y = 10000
                    limit_x = 10000
                if abs(delta_y) < limit_y:
                    if abs(delta_x) < limit_x:
                        pyautogui.move(delta_x, delta_y, duration=0.1)
                        d = delta_y + delta_x
                        if turn == 0 and d < 4:
                            turn = 1
                            win32api.keybd_event(87, 0, 0, 0)
                            time.sleep(0.05)
                            # ticks1 = ticks1 + 1
                            win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
                            d = 10000
                        if turn == 1 and d < 4:
                            if ticks % 5000 == 0:
                                turn = 0
                            else:
                                ticks = ticks + 1
                                print(ticks)
                            win32api.keybd_event(9, 0, 0, 0)
                            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                            win32api.keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
                            time.sleep(0.4)
                            if upordown % 2 != 0:
                                updown()
                            d = 10000
                    else:
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                        time.sleep(0.02)
            elif rotate % 2 != 0:
                # win32api.keybd_event(87, 0, 0, 0)
                # time.sleep(0.1)
                # win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
                d = 10000
                ticks = 0
                if turn == 1:
                    turn = 0
                    win32api.keybd_event(83, 0, 0, 0)
                    # time.sleep((0.1 + 0.1*ticks1))
                    ticks1 = 0
                    win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)
                    add_x = -add_x
                    add_y = -add_y
                    pyautogui.move(add_x, add_y, duration=0.01)
                pyautogui.move(-100, 0, duration=0.05)
                add_x1 = add_x1 - 100
                if add_x1 % (-2400) == 0:    #转一圈就暂停然后向前走一段距离
                    win32api.keybd_event(87, 0, 0, 0)
                    time.sleep(0.05)
                    win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                add_x = 0
                add_y = 0
                time.sleep(0.02)
            else:
                add_x1 = 0
    # if mode == 2:    #砍树模式，与挖矿模式同理，只不过是用的模型不同
    #     choice = choice2
    #     results = model2(screen)
    #     nearest_center = get_closet_center(results.xyxy[0], choice)
    #     current_position = pyautogui.position()
    #     results = np.squeeze(results.render())
    #     results = cv2.cvtColor(results, cv2.COLOR_BGR2RGB)
    #     cv2.imshow("detect", results)
    #     cv2.waitKey(1)
    #     x = rect[0]
    #     y = rect[1]
    #     w = rect[2] - x
    #     h = rect[3] - y
    #     if nearest_center:
    #         delta_x = (nearest_center[0] - current_position[0]) / 2.5
    #         delta_y = (nearest_center[1] - current_position[1]) / 2.5
    #     if (win32api.GetAsyncKeyState(0x04) & 0x8000) > 0:    #鼠标中键启动追踪模式
    #         start = start + 1
    #         time.sleep(0.5)
    #         if start % 2 == 0 and start != 2:
    #             print("link close o.0")
    #             play('otto2.wav')
    #             win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
    #         if start % 2 != 0:
    #             print("link start!")
    #             play('otto.wav')
    #             win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
    #     if start % 2 != 0:
    #         if nearest_center:
    #             pyautogui.move(delta_x, delta_y, duration=0.1)
    #             time.sleep(0.02)
    #             win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    #             time.sleep(0.02)
    #         else :
    #             win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    #             time.sleep(0.02)


    if mode == 3:    #砍怪模式，此模式中，鼠标会自动锁定到检测到的怪身上并不断按下攻击键，以实现自动打怪的目的
        choice = choice3
        results = model3(screen)
        nearest_center = get_closet_center(results.xyxy[0], choice)
        current_position = pyautogui.position()
        results = np.squeeze(results.render())
        results = cv2.cvtColor(results, cv2.COLOR_BGR2RGB)
        cv2.imshow("detect", results)
        cv2.waitKey(1)
        x = rect[0]
        y = rect[1]
        w = rect[2] - x
        h = rect[3] - y
        if (win32api.GetAsyncKeyState(0x04) & 0x8000) > 0:    #鼠标中键启动追踪模式
            start = start + 1
            time.sleep(0.5)
            if start % 2 == 0 and start != 2:
                print("link close o.0")
                play('otto2.wav')
                win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
            if start % 2 != 0:
                print("link start!")
                play('otto.wav')
                win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
        if start % 2 != 0:
            if nearest_center:
                delta_x = (nearest_center[0] - current_position[0]) / 2.5
                delta_y = (nearest_center[1] - current_position[1]) / 2.5
                add_x = add_x + delta_x
                add_y = add_y + delta_y
                pyautogui.move(delta_x, delta_y, duration=0.01)
                time.sleep(0.02)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
            else:
                add_x = -add_x
                add_y = -add_y
                pyautogui.move(add_x, add_y, duration=0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                add_x = 0
                add_y = 0
                time.sleep(0.02)
    if mode == 4:    #射箭模式，MC中的弓箭有抛物线，因此当检测到怪物目标时，当前鼠标会在x轴上牢牢锁在怪物的x坐标上，而y轴可以自由动，以实现控制抛物线精准打击的目的
        choice = choice3
        results = model3(screen)
        nearest_center = get_closet_center(results.xyxy[0], choice)
        current_position = pyautogui.position()
        results = np.squeeze(results.render())
        results = cv2.cvtColor(results, cv2.COLOR_BGR2RGB)
        cv2.imshow("detect", results)
        cv2.waitKey(1)
        x = rect[0]
        y = rect[1]
        w = rect[2] - x
        h = rect[3] - y
        if nearest_center:
            delta_x = (nearest_center[0] - current_position[0]) / 2.5
        if (win32api.GetAsyncKeyState(0x04) & 0x8000) > 0:    #鼠标中键启动追踪模式
            start = start + 1
            time.sleep(0.5)
            if start % 2 == 0 and start != 2:
                print("link close o.0")
                play('otto2.wav')
                win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
            if start % 2 != 0:
                print("link start!")
                play('otto.wav')
                win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
        if start % 2 != 0:
            if nearest_center:
                pyautogui.move(delta_x, 0 , duration=0.1)    #鼠标会在X轴上位移而不在Y轴上面位移
                time.sleep(0.02)
            else:
                time.sleep(0.02)