import sys
import threading
import time
import tkinter
import tkinter.messagebox
import tkinter.ttk
from tkinter import filedialog
from PIL import Image
import keyboard
import psutil
import serial
import serial.tools.list_ports
import win32gui
import win32process
from PIL.Image import Resampling

firstEnter = True

'''
搜索电脑上已安装应用
'''


def searchAppList():
    oriList = []
    pids = psutil.pids()
    for pid in pids:
        oriList.append(psutil.Process(pid).name())
    return list(set(oriList))


'''
搜索已连接上电脑串口
'''


def searchPortsList():
    global selectedApp
    print(selectedApp.get())
    # 以后自定义对话框位置计算，目前未使用
    # root_x = window.winfo_rootx()
    # root_y = window.winfo_rooty()
    # dia_x=root_x+100
    # dia_y=root_y+100
    portList = list(serial.tools.list_ports.comports())
    if len(portList) != 0:
        # 如果串口集合不为0则去判断是否有自己的设备，有的话就连接没有的话就不处理
        firstPort = portList[0]
        print(firstPort)
        return portList
    if not firstEnter:
        tkinter.messagebox.showinfo(title="提示", message='未搜索到设备', parent=window)


window = tkinter.Tk()
selectedApp = tkinter.StringVar()
selectedPort = tkinter.StringVar()
portsList = searchPortsList()
appList = searchAppList()


def connectSerialPort():
    selected_file_name = tkinter.filedialog.askopenfilename()
    image = Image.open(selected_file_name)
    w, h = image.size  # 查看图片尺寸
    print(w, h)
    image.resize((128, 96), Resampling.LANCZOS)
    # 保存再什么地方了？
    image.save('D:\\abc.png')


'''
关闭窗口方法
'''


def closeWindow():
    ans = tkinter.messagebox.askokcancel(title='Warning', message='是否确认关闭应用', parent=window)
    if ans:
        appMoniterThread.kill()
        hotKeyMoniterThread.kill()
        window.destroy()
        sys.exit(0)
    else:
        return


'''
前端应用检测线程,每1秒检测一次
用于不断检测什么应用在前端获取焦点，获取到改变后将设置的数据发送给键盘
'''


class AppMoniterThread(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
        self.killed = False

    def run(self):
        oldAppName = ""
        while not self.killed:
            pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
            if pid[-1] >= 0:
                appName = psutil.Process(pid[-1]).name()
                if oldAppName != appName:
                    oldAppName = appName
                    print('应用改变，开始发送')
                    print(appName)
            time.sleep(1)

    def kill(self):
        self.killed = True


hotKey = ''
hotkeyEntryText = tkinter.StringVar()

'''
热键检测线程
'''


class HotKeyMoniterThread(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
        self.killed = False

    def run(self):
        while not self.killed:
            # 保持要设置为false,否则多次快捷键会粘连
            global hotKey
            hotKey = keyboard.read_hotkey(suppress=False)
            global hotkeyEntryText
            hotkeyEntryText.set(hotKey)
            print(hotKey)
            time.sleep(0.5)

    def kill(self):
        self.killed = True

    def setActive(self):
        self.killed = False


appMoniterThread = AppMoniterThread()
hotKeyMoniterThread = HotKeyMoniterThread()


def handleFocusIn(event):
    global hotKeyMoniterThread
    hotKeyMoniterThread = HotKeyMoniterThread()

    hotKeyMoniterThread.start()


def handleFocusOut(event):
    hotKeyMoniterThread.kill()


'''
窗体线程

创建窗体，同时增加我们所需的布局

'''


def createWindow():
    global firstEnter
    firstEnter = False
    # 监听关闭
    window.protocol('WM_DELETE_WINDOW', closeWindow)
    window.title('touchpad-companion')
    window.geometry('634x504')
    window.resizable(False, False)
    # 初始化窗体居中屏幕
    window_height = 504
    window_width = 634
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    window.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

    pixelVirtual = tkinter.PhotoImage(width=1, height=1)
    leftFrame = tkinter.Frame(window, width="424", height="484", bg="#ffffff", bd="0")
    leftFrame.place(x=10, y=10)
    # 第一行
    button11 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button11.place(x=10, y=20)
    button12 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button12.place(x=148, y=20)
    button13 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button13.place(x=286, y=20)
    # 第二行
    button21 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button21.place(x=10, y=136)
    button22 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button22.place(x=148, y=136)
    button23 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button23.place(x=286, y=136)
    # 第三行
    button31 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button31.place(x=10, y=252)
    button32 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button32.place(x=148, y=252)
    button33 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button33.place(x=286, y=252)
    # 第四行
    button31 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button31.place(x=10, y=368)
    button32 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button32.place(x=148, y=368)
    button33 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button33.place(x=286, y=368)

    '''
    右侧功能选项区
    '''
    rightFrame = tkinter.Frame(window, width="170", height="484", bg="#ffffff", bd="0")
    rightFrame.place(x=454, y=10)
    # 应用选择
    portLabel = tkinter.Label(rightFrame, text="关联应用", bg="#ffffff", fg="#606266")
    portLabel.place(x=10, y=10)
    global selectedApp
    combobox = tkinter.ttk.Combobox(
        master=rightFrame,  # 父容器
        height=8,  # 高度,下拉显示的条目数量
        width=16,  # 宽度
        state='readonly',
        font=('', 12),
        textvariable=selectedApp,
        values=appList,
    )
    combobox.place(x=10, y=35)
    # button = tkinter.Button(rightFrame, text="刷新", activebackground="#53abff", activeforeground="#ffffff",
    #                         bd='0', fg="#ffffff", bg="#409eff", command=searchAppList)
    # button.place(x=130, y=35)
    # 快捷键输入
    portLabel = tkinter.Label(rightFrame, text="快捷键", bg="#ffffff", fg="#606266")
    portLabel.place(x=10, y=70)

    hotkeyEntry = tkinter.Entry(rightFrame, bg="#ffffff", textvariable=hotkeyEntryText)
    hotkeyEntry.place(x=10, y=90)
    hotkeyEntry.bind("<FocusIn>", handleFocusIn)
    hotkeyEntry.bind("<FocusOut>", handleFocusOut)
    # 串口选择
    portLabel = tkinter.Label(rightFrame, text="通信串口", bg="#ffffff", fg="#606266")
    portLabel.place(x=10, y=425)
    global selectedPort
    selectedPort = tkinter.StringVar()
    # 设置选中
    # selectedPort.set('COM2')
    combobox = tkinter.ttk.Combobox(
        master=rightFrame,  # 父容器
        height=8,  # 高度,下拉显示的条目数量
        width=8,  # 宽度
        state='readonly',
        font=('', 16),
        textvariable=selectedPort,
        values=portsList,
    )
    combobox.place(x=10, y=450)
    button = tkinter.Button(rightFrame, text="刷新", activebackground="#53abff", activeforeground="#ffffff",
                            bd='0', fg="#ffffff", bg="#409eff", command=searchPortsList)
    button.place(x=130, y=450)
    window.mainloop()


if __name__ == '__main__':
    appMoniterThread.start()
    createWindow()
