import sys
import threading
import time
import tkinter
import tkinter.ttk
import tkinter.messagebox
import winreg
import psutil
import win32gui
import win32process
import serial
import serial.tools.list_ports

# 定义检测位置
sub_key = [r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
           r'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall']
window = tkinter.Tk()


def searchAppList():
    oriList = []
    pids = psutil.pids()
    for pid in pids:
        oriList.append(psutil.Process(pid).name())
    return list(set(oriList))


appList = searchAppList()


def searchPortsList():
    portList = list(serial.tools.list_ports.comports())
    if len(portList) != 0:
        print(len(portList))
        return portList
    tkinter.messagebox.showinfo(title="提示", message='未搜索到设备', parent=window)


portsList = searchPortsList()


def connectSerialPort():
    print('hello')


def closeWindow():
    ans = tkinter.messagebox.askokcancel(title='Warning', message='是否确认关闭应用')
    if ans:
        window.destroy()
        appMoniterThread.kill()
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


appMoniterThread = AppMoniterThread()


def createWindow():
    # 监听关闭
    window.protocol('WM_DELETE_WINDOW', closeWindow)
    window.title('touchpad-companion')
    window.geometry('580x520')
    window.resizable(False, False)
    pixelVirtual = tkinter.PhotoImage(width=1, height=1)
    leftFrame = tkinter.Frame(window, width="380", height="500", bg="#ffffff", bd="0")
    leftFrame.place(x=10, y=10)
    # 第一行
    button11 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button11.place(x=10, y=10)
    button12 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="128", height="96",
                              command=connectSerialPort)
    button12.place(x=148, y=20)
    button13 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button13.place(x=260, y=20)
    # 第二行
    button21 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button21.place(x=20, y=140)
    button22 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button22.place(x=140, y=140)
    button23 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button23.place(x=260, y=140)
    # 第三行
    button31 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button31.place(x=20, y=260)
    button32 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button32.place(x=140, y=260)
    button33 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button33.place(x=260, y=260)
    # 第四行
    button31 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button31.place(x=20, y=380)
    button32 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button32.place(x=140, y=380)
    button33 = tkinter.Button(leftFrame, activebackground="#c8c9cc", bd='0', image=pixelVirtual, compound="c",
                              width="100", height="100",
                              command=connectSerialPort)
    button33.place(x=260, y=380)

    '''
    右侧功能选项区
    '''
    rightFrame = tkinter.Frame(window, width="170", height="500", bg="#ffffff", bd="0")
    rightFrame.place(x=400, y=10)
    # 应用选择
    portLabel = tkinter.Label(rightFrame, text="关联应用", bg="#ffffff", fg="#606266")
    portLabel.place(x=10, y=10)
    selectedApp = tkinter.StringVar()
    combobox = tkinter.ttk.Combobox(
        master=rightFrame,  # 父容器
        height=8,  # 高度,下拉显示的条目数量
        width=12,  # 宽度
        state='readonly',
        font=('', 14),
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
    hotkeyEntry = tkinter.Entry(rightFrame, bg="#ffffff")
    hotkeyEntry.place(x=10, y=90)
    # 串口选择
    portLabel = tkinter.Label(rightFrame, text="通信串口", bg="#ffffff", fg="#606266")
    portLabel.place(x=10, y=435)
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
    combobox.place(x=10, y=460)
    button = tkinter.Button(rightFrame, text="刷新", activebackground="#53abff", activeforeground="#ffffff",
                            bd='0', fg="#ffffff", bg="#409eff", command=searchPortsList)
    button.place(x=130, y=460)
    window.mainloop()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    appMoniterThread.start()
    createWindow()
