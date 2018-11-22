# -*- coding: utf-8 -*-
# http://www.snb-vba.eu/VBA_Outlook_external_en.html
import win32com.client as win32
import warnings
import datetime
import pandas as pd
import os
import shutil
import subprocess
import pyautogui
import time
import sys


# open file
file_usr = open("./usrinfo.txt", "a+")
# read file
user_read = file_usr.read()
# define user and password
[user_id, user_pw] = user_read.split(",")
# close file
file_usr.close()
# used for log comment
time_now = datetime.datetime.now().strftime('%m/%d-%H:%M')
# ignore warn
warnings.filterwarnings('ignore')
# 控制函数每隔1秒执行下一个函数
pyautogui.PAUSE = 1
# 启动自动防故障功能
pyautogui.FAILSAFE = True
# open file
script_path = r"F:\Scripting\Export\SHA_TSF\Air_Export_TSF-1.eds"


def open_log():
    log = open("./log.txt", "a+")
    return log


if os.path.exists(r'c:\Script\TYT'):
    root_path = r"c:\Script\TYT"
else:
    os.makedirs(r"c:\Script\TYT")
    root_path = r"c:\Script\TYT"


if os.path.exists(r'c:\Script\TYT\datafile'):
    data_path = r"c:\Script\TYT\datafile"
else:
    os.makedirs(r"c:\Script\TYT\datafile")
    data_path = r"c:\Script\TYT\datafile"


if os.path.exists(r'c:\script\TYT\input'):
    input_path = r"c:\script\TYT\input"
else:
    os.makedirs(r"c:\script\TYT\input")
    input_path = r"c:\script\TYT\input"


if os.path.exists(r'c:\script\TYT\output'):
    output_path = r"c:\script\TYT\output"
else:
    os.makedirs(r"c:\script\TYT\output")
    output_path = r"c:\script\TYT\output"


if os.path.exists(r'c:\script\TYT\history'):
    history_path = r"c:\script\TYT\history"
else:
    os.makedirs(r"c:\script\TYT\history")
    history_path = r"c:\script\TYT\history"


def send_email():
    sub = 'TSF Script Running Failed'
    body = 'Computer IP: 192.8.12.178\r\nScriptPath: pvg-fs\sys\Scripting\Export\SHA_TSF\Air_Export_TSF.eds'
    outlook = win32.Dispatch('outlook.application')
    receivers = ['greg.he@expeditors.com']
    mail = outlook.CreateItem(0)
    mail.To = receivers[0]
    mail.Subject = sub.decode('utf-8')
    mail.Body = body.decode('utf-8')
    mail.Send()
    time.sleep(1)


class GetUserinfo:

    def __init__(self, user, passwd):
        self.user = user
        self.passwd = passwd

    def get_userid(self):
        self.user = raw_input('Login Username:')
        while self.user[0:4] != "SHA-" and self.user[0:4] != "sha-":
            print (unicode("Username Input Invalid, please try again...", encoding='utf-8'))
            self.user = raw_input('Login Username:')
        else:
            file_usr.write(self.user + ",")
            return self.user

    def get_passwd(self):
        self.passwd = raw_input('login Password:')
        while len(self.passwd) < 8:
            print (unicode("Not meet the requirement of password policy, please input at least 8 bites password...", encoding='utf-8'))
            self.passwd = raw_input('login Password:')
        else:
            file_usr.seek(0, 1)
            file_usr.write(self.passwd)


SHA = GetUserinfo("self.userid", "self.userpw")


if (len(user_read) == 0) or ("," not in user_read) or (user_read.find(',') == len(user_read) - 1):
    print (unicode("\nPlease input your etms system username and password, username for example: sha-expeditors", encoding='utf-8'))
    file_usr = open("./usrinfo.txt", "w+")
    SHA.get_userid()
    SHA.get_passwd()


if (len(user_read) != 0) and ("," in user_read) and (user_read.find(',') != len(user_read) - 1):
    if user_id[0:3] == "sha":
        print (unicode("Your Username:", encoding='utf-8'))+user_id
        print (unicode("Your Password:", encoding='utf-8'))+user_pw


def close_app():
    wmi = win32.GetObject('winmgmts:')
    process_desktop = wmi.ExecQuery('select * from Win32_Process where Name="%s"' % "desktop.exe")
    if len(process_desktop) > 0:
        subprocess.call("taskkill /f /im desktop.exe")
        time.sleep(2)
    process_outlook = wmi.ExecQuery('select * from Win32_Process where Name="%s"' % "outlook.exe")
    if len(process_outlook) > 0:
        pass
    else:
        subprocess.Popen(r'C:\Program Files (x86)\Microsoft Office\Office16\outlook.exe')
        time.sleep(20)


def check_network():
    open_log().write("Network")
    open_log().close()
    net_check1 = subprocess.call(['ping', "10.0.2.23"])
    net_check2 = subprocess.call(['ping', "192.8.12.45"])
    if net_check1 is 0 and net_check2 is 0:
        open_log().write(":OK---")
        open_log().close()
    else:
        open_log().write(":Failed\n")
        open_log().close()
        os.system("shutdown /r /t 3")
        sys.exit(1)


def map_drive():
    # initialize
    drive_letter = 'F:'
    network_path = '\\\\pvg-fs\\sys'
    user = user_id
    password = user_pw
    # Disconnect anything on drive letter Q
    win_cmd = 'NET USE ' + drive_letter + ' /delete'
    subprocess.call(win_cmd, shell=True)
    # Connect to map network drive to letter Q
    win_cmd = 'NET USE ' + drive_letter + ' ' + network_path + ' /User:' + user + ' ' + password
    subprocess.call(win_cmd, shell=True)


def save_attch(subject, attach):
    # print
    print(unicode("\nSearching...", encoding='utf-8'))
    # Connect with MS Outlook
    outlook = win32.Dispatch('outlook.application').GetNamespace("MAPI")
    # connect to Inbox Items, "6" refers to the inobx item of a folder
    inbox_msgs = outlook.GetDefaultFolder(6).Items
    # get the email/s
    msg = inbox_msgs.GetLast()
    # define a sub folder
    sub_folder = outlook.GetDefaultFolder(6).folders("TSFR")
    # set n equal to 0
    n = 0
    # start searching email
    while msg:
        # filter subject
        if msg.Subject.startswith(unicode(subject, encoding='utf-8')):
            # get attachment
            attachments = msg.Attachments
            # get the only one attachment
            attachment = attachments.Item(1)
            # save attachment on local
            attachment.SaveAsFile(data_path + '\\' + (unicode(attach, encoding='utf-8')))
            # sum n
            n += 1
            # move the received email to sub folder
            msg.move(sub_folder)
            # if attach file saved, break
            if n == 1:
                print(unicode("\nSavedFile:%d..." % n, encoding='utf-8'))
                # return result
                return 0
        # get the next email
        msg = inbox_msgs.GetPrevious()
    # time delay
    time.sleep(1)


def data_process(xlsx, csv):
    # load excel file and save as a new csv file
    pd.read_excel(os.path.join(data_path, xlsx)).to_csv(os.path.join(input_path, "abuse.csv"), encoding='utf-8')
    # read input_csv file
    dfc = pd.read_csv(os.path.join(input_path, "abuse.csv"), encoding='utf-8')
    # split date and time columns
    dfc['MBL'] = dfc[u'天运通●日常业务统计报表'].str[0:3] + "-" + dfc[u'天运通●日常业务统计报表'].str[3:]
    dfc['Event_Date'] = dfc['Unnamed: 18'].str.split(' ').str[0].str[0:4] + dfc['Unnamed: 18'].str.split(' ').str[0].str[5:7] + dfc['Unnamed: 18'].str.split(' ').str[0].str[8:10]
    dfc['Event_Time'] = dfc['Unnamed: 18'].str.split(' ').str[1].str[0:5]
    dfc['Remarks'] = dfc['Unnamed: 5']
    # filter and save as input csv file
    dfc.loc[:, ['MBL', 'Event_Date', 'Event_Time', 'Remarks']].to_csv(os.path.join(input_path, csv), encoding='utf-8', index=None)
    # loading input csv file
    dfci = pd.read_csv(os.path.join(input_path, csv))
    # remove duplicates
    dfci.drop_duplicates(subset=['MBL', 'Event_Date', 'Event_Time', 'Remarks'], keep='first', inplace=True)
    # delete Null rows
    dfci.dropna(axis=0, how='any', inplace=True)
    # save data
    dfci.to_csv(os.path.join(input_path, csv), header=False, encoding='utf-8', index=None)
    # delete abuse csv file
    os.remove(os.path.join(input_path, "abuse.csv"))
    # remove to done folder
    shutil.move(os.path.join(data_path, xlsx), os.path.join(history_path, xlsx))
    # format date
    time_xls = datetime.datetime.now().strftime('%m-%d-%H-%M-%S')
    # rename
    os.rename(os.path.join(history_path, xlsx), os.path.join(history_path, str(time_xls) + ".xlsx"))
    # print
    print(unicode("\nFormatData...", encoding='utf-8'))
    # add time delay variable
    time.sleep(1)
    # write log
    open_log().write("FormatData:OK---")
    # close log
    open_log().close()
    # time delay
    time.sleep(1)


def start_app():
    print(unicode("\nStartApp...", encoding='utf-8'))
    open_log().write("StartApp")
    subprocess.Popen([r'C:\Program Files (x86)\Expeditors\Desktop\Desktop.exe'])
    wmi = win32.GetObject('winmgmts:')
    process = wmi.ExecQuery('select * from Win32_Process where Name="%s"' % "desktop.exe")
    if len(process) in range(1, 10):
        time.sleep(1)
        open_log().write(":OK---")
        open_log().close()
        return 0
    elif len(process) > 9:
        open_log().write(":Process>9---RestartComputer\n")
        open_log().close()
        os.system("shutdown /r /t 1")
        sys.exit(1)
    else:
        open_log().write(":Process=0---RestartComputer\n")
        open_log().close()
        os.system("shutdown /r /t 1")
        sys.exit(1)


def login_app():
    print(unicode("\nLoginApp...", encoding='utf-8'))
    open_log().write("Login")
    open_log().close()
    for x in range(20):
        time.sleep(1)
        pyautogui.click()
        username_image = pyautogui.locateOnScreen('./png/username.png', grayscale=True)
        if username_image is not None:
            break
    username_image = pyautogui.locateOnScreen('./png/username.png', grayscale=True)
    if username_image is not None:
        ux, uy = pyautogui.center(username_image)
        pyautogui.moveTo(ux, uy)
        pyautogui.click()
        pyautogui.moveTo(ux+179, uy)
        pyautogui.click()
        pyautogui.click()
        pyautogui.moveTo(ux+90, uy-5)
        pyautogui.click()
        pyautogui.press('shift')
        pyautogui.typewrite(user_id)
        pyautogui.moveTo(ux+90, uy+25)
        pyautogui.click()
        pyautogui.click()
        pyautogui.typewrite(user_pw)
        pyautogui.press('enter')
        open_log().write(":OK---")
        open_log().close()
        time.sleep(1)
        return 0
    else:
        open_log().write(":Failed\n")
        open_log().close()
        send_email()
        close_app()
        sys.exit(1)


def click_lunch():
    print(unicode("\nClickLunch...", encoding='utf-8'))
    open_log().write("Lunch")
    for x in range(20):
        time.sleep(1)
        pyautogui.click()
        launch = pyautogui.locateOnScreen(image="./png/launch.png", grayscale=True)
        if launch is not None:
            break
    launch = pyautogui.locateOnScreen(image="./png/launch.png", grayscale=True)
    if launch is not None:
        xl, yl = pyautogui.center(launch)
        pyautogui.moveTo(xl, yl)
        pyautogui.click()
        open_log().write(":OK---")
        open_log().close()
        time.sleep(1)
        return 0
    else:
        open_log().write(":Failed\n")
        open_log().close()
        send_email()
        close_app()
        sys.exit(1)


def click_export():
    print(unicode("\nClickExport...", encoding='utf-8'))
    open_log().write("Export")
    for i in range(20):
        export = pyautogui.locateOnScreen(image="./png/export-1920.png", grayscale=True)
        time.sleep(1)
        if export is not None:
            break
    export = pyautogui.locateOnScreen(image="./png/export-1920.png", grayscale=True)
    if export is not None:
        ex, ey = pyautogui.center(export)
        pyautogui.moveTo(ex, ey)
        time.sleep(0.5)
        pyautogui.click()
        open_log().write(":OK---")
        open_log().close()
        time.sleep(1)
        return 0
    else:
        open_log().write(":Failed\n")
        open_log().close()
        send_email()
        close_app()
        sys.exit(1)


def command_input():
    print(unicode("\nClickCommand...", encoding='utf-8'))
    open_log().write("Screen")
    for x in range(20):
        time.sleep(1)
        command = pyautogui.locateOnScreen(image="./png/command-1920.png", grayscale=True)
        if command is not None:
            break
    command = pyautogui.locateOnScreen(image="./png/command-1920.png", grayscale=True)
    if command is not None:
        cx, cy = pyautogui.center(command)
        pyautogui.moveTo(cx, cy+50)
        pyautogui.click()
        pyautogui.press('backspace')
        pyautogui.press('backspace')
        pyautogui.typewrite("HIST")
        pyautogui.press('f8')
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(2)
        open_log().write(":OK---")
        open_log().close()
        time.sleep(1)
        return 0
    else:
        open_log().write(":Failed\n")
        open_log().close()
        send_email()
        close_app()
        sys.exit(1)


def run_script(path):
    print(unicode("\nRunScript...", encoding='utf-8'))
    open_log().write("RunScript")
    # circle
    for x in range(20):
        time.sleep(1)
        pyautogui.click()
        script = pyautogui.locateOnScreen(image="./png/script-1920.png", grayscale=True)
        if script is not None:
            break
    script = pyautogui.locateOnScreen(image="./png/script-1920.png", grayscale=True)
    if script is not None:
        sx, sy = pyautogui.center(script)
        pyautogui.moveTo(sx-300, sy)
        pyautogui.typewrite(path)
        pyautogui.press('shift')
        pyautogui.press('enter')
        open_log().write(":OK---")
        open_log().close()
        time.sleep(2)
        return 0
    else:
        open_log().write(":Failed\n")
        open_log().close()
        send_email()
        close_app()
        sys.exit(1)


def check_result():
    open_log().write("CheckResult")
    open_log().close()
    print(unicode("\nCheckResult...", encoding='utf-8'))
    # circle
    for x in range(2000):
        time.sleep(1)
        status = pyautogui.locateOnScreen(image="./png/complete-1920.png", grayscale=True)
        if status is not None:
            break
    status = pyautogui.locateOnScreen(image="./png/complete-1920.png", grayscale=True)
    if status is not None:
        stx, sty = pyautogui.center(status)
        pyautogui.moveTo(stx, sty+50)
        pyautogui.click()
        pyautogui.click()
        close_app()
        time.sleep(1)
        shutil.move(os.path.join(input_path, "input.csv"), os.path.join(history_path, "input.csv"))
        time_input = datetime.datetime.now().strftime('%m-%d-%H-%M-%S')
        time.sleep(1)
        os.rename(os.path.join(history_path, "input.csv"), os.path.join(history_path, str(time_input) + ".csv"))
        open_log().write(":OK\n")
        open_log().close()
        time.sleep(2)
        return 0
    else:
        open_log().write(":Failed\n")
        open_log().close()
        send_email()
        close_app()
        sys.exit(1)


def trigger():
    while True:
        if os.path.exists(os.path.join(input_path, "input.csv")):
            open_log().write(time_now + " ")
            map_drive()
            close_app()
            check_network()
            open_log().write("FormatData:OK---")
            open_log().close()
            all_in_one()
        elif save_attch("天运通●日常业务统计报表", "tyt.xlsx") is 0:
            open_log().write(time_now + " ")
            map_drive()
            close_app()
            check_network()
            open_log().close()
            data_process("tyt.xlsx", "input.csv")
            all_in_one()
        else:
            close_app()
            sys.exit(0)


def all_in_one():
    pyautogui.hotkey('winleft', 'd')
    if start_app() is 0:
        if login_app() is 0:
            if click_lunch() is 0:
                if click_export() is 0:
                    if command_input() is 0:
                        if run_script(script_path) is 0:
                            while check_result():
                                trigger()


if __name__ == "__main__":
    trigger()
