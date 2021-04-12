# coding=utf-8
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
import pyHook
import pythoncom
import xlrd
import xlwt
import pyperclip
from pynput import mouse, keyboard
import threading
import sys
import re
from openpyxl import Workbook, load_workbook


def copy():
    k.press_key(k.control_l_key)
    k.tap_key("c")  # 改小写！！！！ 大写的话由于单进程会触发shift键 ctrl键就失效了
    k.release_key(k.control_l_key)




def getCopy(noresult=None,maxTime=1.3,isDone=False):
    # maxTime = 3  # 3秒复制 调用copy() 不管结果对错
    if(maxTime<=0 or isDone):
        return noresult
    pyperclip.copy('')
    time.sleep(0.3)
    # print('doing')
    copy()
    result = pyperclip.paste()
    if(result==''):
        return getCopy(noresult,maxTime-0.3,False)
    else:
        return getCopy(result,maxTime-0.3,True)

    #print('debug:'+str(result))
    return result





def tapkey(key, count=1, waitTime=0.1):
    for i in range(0, count):
        k.tap_key(key)
        time.sleep(waitTime)


def Quit():
    global end
    end = True


def Do():
    global index_count
    global start
    if start:
        time.sleep(1)
        # 判断编号是否为12位数字
        tapkey(k.escape_key)

        tapkey(k.left_key, 8)
        tapkey(k.down_key)
        tapkey(k.right_key, 3)
        tapkey(k.up_key)
        code = getCopy()  # 项目编号
        # print(code)
        isRightNum = False
        index_count = index_count + 1
        length = len(str(code))
        if (str(code).isdigit()):
            if (length == 12):
                # print('is 12 number')
                isRightNum = True
            else:
                print('is not 12 number')
        elif (length == 6):
            pattern = re.compile('\d\dB\d+')
            match = pattern.match(str(code))
            if (match):
                isRightNum = True
                print('特殊id  1')
            else:
                print('不对的6位数')
        elif(length==9):#03B050001 03B020001
            pattern = re.compile('\d\dB\d+')
            match = pattern.match(str(code))
            if (match):
                isRightNum = True
                print('特殊id  1')
            else:
                print('不对的9位数')
        elif (length == 13):
            pattern = re.compile('Z\d+')
            match = pattern.match(str(code))
            if (match):
                isRightNum = True
                print('特殊id  2')
            else:
                print('不对的13位数')
        else:
            print('is NOT number')
        if ('-' in str(code)):
            isRightNum = False
        if (isRightNum == False):
            print(code)
            Quit()
        tapkey(k.down_key)
        tapkey(k.escape_key)
        tapkey(k.right_key, 2)
        tapkey(k.up_key)
        # 主代码---------------
        projectCharactor = getCopy()  # 项目特征
        downNumber = 0
        while (downNumber < 14):
            tapkey(k.down_key)
            nowProjectCharactor = getCopy()
            downNumber = downNumber + 1
            if (downNumber < 4):
                if (nowProjectCharactor == None):
                    continue
                else:
                    print('不处理_index：' + str(index_count))
                    print('项目特征：' + projectCharactor)

                    print('<4')
                    return
            # elif(downNumber==3):
            #     #三个头的 改Q的
            elif (downNumber == 4):
                if (nowProjectCharactor == None):
                    continue
                elif (nowProjectCharactor != projectCharactor):
                    projectCharactor = nowProjectCharactor
                    # print('4')
                    break
            elif (downNumber > 4):
                print('大于4的项目')
                if (nowProjectCharactor == None):
                    continue
                else:
                    print('不处理_index：' + str(index_count))
                    print('项目特征：' + projectCharactor)

                    return
        else:
            print('>=13')
            Quit()
            return
        print('开始检测综合单价')
        tapkey(k.up_key, 3)
        tapkey(k.escape_key)
        tapkey(k.right_key, 1)

        #根据单位下浮------------------------------
        tapkey(k.up_key)
        unit=getCopy()
        tapkey(k.down_key)
        tapkey(k.escape_key)

        tapkey(k.right_key, 3)
        tapkey(k.up_key)

        zong = float(getCopy(0))
        tapkey(k.down_key)
        tapkey(k.escape_key)
        tapkey(k.right_key)
        tapkey(k.up_key)
        kong = float(getCopy())
        swimValue = abs((kong) - (zong)) / (kong)

        lowest = kong * 0.85

        #lowest = kong * 0.8##智能化--------------------------------------------

        arg = 1
        UnitCloseArg='m' in unit or 'kg' in unit
        if(UnitCloseArg):
            if(lowest>1000):
                arg=0.5
            elif(lowest>500):
                arg=0.3
            if(lowest>100):
                arg=0.25
            elif(lowest>0):
                arg=0.2
        else:
            if (lowest > 10000):
                arg = 100
            elif(lowest>5000):
                arg=50
            elif (lowest > 3000):
                arg = 30
            elif (lowest > 1000):
                arg = 20
            elif (lowest > 500):
                arg = 15
            elif (lowest > 100):
                arg = 10
            elif (lowest > 50):
                arg = 5
            elif (lowest > 30):
                arg = 2
            elif (lowest > 10):
                arg = 1
            elif (lowest > 0):
                arg = 0.2
            else:
                print('负数')



        condition1 = zong <= lowest - 0.01
        condition2 = zong > lowest + arg + 0.01

        # for lowest


        if (condition2 or condition1):
            # print(swimValue)
            # print(zong)
            # print(kong)
            tapkey(k.escape_key)
            tapkey(k.left_key)
            tapkey(k.down_key, 2)

            nowValue=float(getCopy(0))



            value = ''


            if (nowValue == None):
                print('空空空')
                # print(projectCharactor)
                if (zong > lowest):
                    print('无主材价，综合>控制,减不动')
                else:
                    value = float(abs((lowest) - (zong))) + arg
                    # print(value)
            else:

                if (zong > lowest + arg):  # TODDODODODODODO
                    minus = (zong - lowest - arg)
                    if (minus > 0):
                        # value=nowValue-minus/1.15#TODDODODODODODO
                        value = nowValue - minus

                    else:
                        print('材料价太少，少到最低价都不够减')
                else:
                    if (zong < lowest):
                        add = lowest + arg - zong
                        # value = nowValue +add / 1.15#TODDODODODODODO
                        value = nowValue + add

                    else:
                        print('满足条件')
            # 比最低价高一点 end


            if (value == ''):
                print('不处理')
                tapkey(k.escape_key)
                tapkey(k.left_key, 4)
                tapkey(k.down_key, 2)
            else:
                # if(value>80):
                # value=int(value)  #不取整了
                if(value<0):
                    value=0
                k.type_string(str(value))
                print(value)
                tapkey(k.enter_key)

                time.sleep(2)

                # 工程量系数 1.03 1.01
                tapkey(k.escape_key)
                tapkey(k.right_key, 6)
                tapkey(k.up_key, 3)

                changed_zong = float(getCopy())

                # temppppppppppppppppppppppppppppppppppppppp
                minus_zong = changed_zong - zong  # 合价
                if (nowValue == None):
                    minus_value = value - 0
                else:
                    minus_value = value - nowValue
                if(minus_value==0):
                    changed_arg=1
                else:
                    changed_arg = minus_zong / minus_value

                print(changed_arg)
                changed_value = value
                if (changed_arg != 1 and changed_arg != 0):

                    if (nowValue == None):
                        print('空空空22')
                        # print(projectCharactor)
                        if (zong > lowest):
                            print('无主材价，综合>控制,减不动22')
                        else:
                            changed_value = (float(abs((lowest) - (zong))) + arg) / changed_arg



                            # print(value)
                    else:

                        if (zong > lowest + arg):  # TODDODODODODODO
                            minus = (zong - lowest) - arg
                            if (minus > 0):
                                # value=nowValue-minus/1.15#TODDODODODODODO
                                changed_value = nowValue - minus / changed_arg

                            else:
                                print('材料价太少，少到最低价都不够减22')

                        else:
                            if (zong < lowest):
                                add = lowest + arg - zong
                                # value = nowValue +add / 1.15#TODDODODODODODO
                                changed_value = nowValue + add / changed_arg

                            else:
                                print('满足条件22')
                    # temp end

                    if (changed_value < 0):
                        print('负数 价格 wrong index::' + str(index_count))
                        print('负数 价格::' + str(changed_value))

                        changed_value = 0

                    tapkey(k.down_key, 2)
                    k.type_string(str(changed_value))
                    # print(changed_arg)
                    print('改系数')
                    print(changed_value)
                    tapkey(k.enter_key)
                    time.sleep(2)

                # 工程量系数 1.03 1.01  end
                else:
                    tapkey(k.down_key, 2)
                    tapkey(k.enter_key)

                # 下一个
                tapkey(k.escape_key)
                tapkey(k.right_key, 2)
                tapkey(k.down_key)

        else:
            print('满足条件，不处理')
            tapkey(k.escape_key)
            tapkey(k.left_key, 5)
            tapkey(k.down_key, 4)

        return









# 我的代码
def onpressed(Key):
    while True:
        # print(Key)
        if (Key == keyboard.Key.caps_lock):  # 开始
            global start
            if (start == True):
                start = False
                print('stop')
            else:
                start = True
                print('go')
        if (Key == keyboard.Key.f3):  # 结束
            sys.exit()

        global end
        if (end):
            sys.exit()
        return True


def main():
    while True:
        # 主程序在这
        Do()


if __name__ == '__main__':
    k = PyKeyboard()
    m = PyMouse()
    end = False
    start = False
    index_count = 0
    # excelUrl = r"C:\Users\Administrator\Desktop\Xing.xlsx"#to do-------------
    # excelUrl = r"C:\Users\123\Desktop\广联达\安装\Xing.xlsx"  # to do-------------


    threads = []
    t2 = threading.Thread(target=main, args=())
    threads.append(t2)
    for t in threads:
        t.setDaemon(True)
        t.start()
    print('press Capital to start')
    print('做之前先把长度为4的，不需要程序进行处理的，排除！！！！')
    #print('智能化0.8看下，164行！！！！')

    with keyboard.Listener(on_press=onpressed) as listener:
        listener.join()

