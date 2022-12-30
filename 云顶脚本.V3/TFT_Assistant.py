import pyautogui
import time
import xlrd
import pyperclip
import random
import cv2
import numpy
from PIL import ImageGrab

# 一些全局变量
from PIL.Image import Image

X_START = 240
X_END = 1700
Y_START = 200
Y_END = 1080
size = (X_START, Y_START, X_END, Y_END)     #截图区域？


#定义鼠标事件

#pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("未找到匹配图片,1秒后重试"+img)
            time.sleep(1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9) #90%相似？
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            time.sleep(1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("重复"+img)
                i += 1
            time.sleep(1)
    elif reTry == -2 :
        image = cv2.data
        if img == "goldbool.png" :
            image = goldbool
        elif img == "bluebool.png" :
            image = bluebool
        # elif img == "gerybool.png":
        #     image = gerybool

            # 使用OpenCV检测，提高精度
            #先截图
            pic: Image = ImageGrab.grab(size)
            # pic.save("target.png")      #保存的图片放哪了？D:\练习\Python\practice\云顶之弈脚本\云顶脚本.V3\
            # target = cv2.imread("target.png")
#             PIL.Image转换成OpenCV格式
            target = cv2.cvtColor(numpy.asarray(pic), cv2.COLOR_RGB2BGR)
            # pic.show()
            #读取图片
            # image = cv2.imread(img)

            #识别
            theight, twidth = image.shape[:2]
            result = cv2.matchTemplate(target, image, cv2.TM_SQDIFF_NORMED)     #100%识别？
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            threashRight = [0.87]   #什么意思？      能够识别了，但是没目标的时候乱识别
            if max_val > threashRight[0]:
                pyautogui.moveTo(min_loc[0] + twidth // 2 + X_START, min_loc[1] + theight // 2 + Y_START)
                pyautogui.mouseDown(button='right')
                pyautogui.mouseUp(button='right')
                #time.sleep(random.randint(1,3))
        else :
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)  # 90%相似？
            if location is not None:

                #pyautogui.mouseUp(button='right', x=100, y=200) # 移动到(100, 200)位置，然后松开鼠标右键
                pyautogui.moveTo(location.x,location.y)
                pyautogui.mouseDown()
                pyautogui.mouseUp()
                if lOrR == "right":
                    pyautogui.mouseDown(button='right') # 按下鼠标右键
                    pyautogui.mouseUp(button='right')
                #pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.5, button=lOrR)
                if img == "upgrade.png" or img == "refresh.png" :
                     time.sleep(random.randint(1,20))
    
        time.sleep(0.1) #随机时间random.random()
        print("没找到{0}，执行下一个操作".format(img))





# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    #行数检查
    if sheet1.nrows<2:  #只有一行
        print("没数据啊哥")
        checkCmd = False
    #每行数据检查
    i = 1
    while i < sheet1.nrows: #当前行小于总行数
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print('第',i+1,"行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd

#任务     按顺序工作？对啊
def mainWork(sheet1):
    i = 1
    while i < sheet1.nrows:
        #取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value  #获取第三行的重复次数
            mouseClick(1,"left",img,reTry)
            #mouseClick(1,"right",img,reTry)
            print("单击左键",img)
        #2代表双击左键
        elif cmdType.value == 2.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("双击左键",img)
        #3代表右键
        elif cmdType.value == 3.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("右键",img) 
        #4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("输入:",inputValue)                                        
        #5代表等待
        elif cmdType.value == 5.0:
            #取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待",waitTime,"秒")
        #6代表滚轮
        elif cmdType.value == 6.0:
            #取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动",int(scroll),"距离")                      
        i += 1

if __name__ == '__main__':
    file = 'cmd.xls'
    #打开文件
    wb = xlrd.open_workbook(filename=file)
    #通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('欢迎使用云顶之弈RPA~')
    #数据检查
    checkCmd = dataCheck(sheet1)
    goldbool = cv2.imread("goldbool.png")
    bluebool = cv2.imread("bluebool.png")
    # gerybool = cv2.imread("greybool.png")
    if checkCmd:
        # key=input('选择功能: 1.做一次 2.一直循环 \n')
        # if key=='1':
        #     #循环拿出每一行指令
        #     mainWork(sheet1)
        # elif key=='2':
            while True:
                # # 执行一轮操作直接一次图，先截图 不准！
                # pic = ImageGrab.grab(size)
                # pic.save("target.png")  # 保存的图片放哪了？D:\练习\Python\practice\云顶之弈脚本\云顶脚本.V3\
                # target = cv2.imread("target.png")
                # # pic.show()

                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")    
    else:
        print('输入有误或者已经退出!')
