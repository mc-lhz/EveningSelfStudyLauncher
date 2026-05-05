import datetime
import time
import os
import subprocess
import threading
import winreg
import psutil
from pynput.keyboard import Controller, Key

def turn_time(hour, minute, second):
    return hour * 60 * 60 + minute * 60 + second

def check_process_exist(process_name):
    for proc in psutil.process_iter(attrs=["cmdline"]):
        try:
            cmd = proc.info["cmdline"]
            if cmd and process_name in " ".join(cmd).lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def press_button(button):
    kb = Controller()
    kb.press(button)
    kb.release(button)

def get_now_time():
    now = datetime.datetime.now()
    hour = now.hour
    minute = now.minute
    second = now.second
    return turn_time(hour, minute, second)

def time_check(wpp_path,ppt_path,start_time,end_time):
    print("时间监控线程运行中")
    # 修复：外层循环添加休眠，防止CPU占满
    while True:
        now_time = get_now_time()
        # 到达时间段
        if start_time <= now_time <= end_time:
            print("到达指定时间，准备启动PPT")
            subprocess.run('''mshta vbscript:CreateObject("Wscript.Shell").popup("即将启动晚自习",5,"WPS",64)(window.close)''')
            subprocess.Popen([wpp_path, ppt_path])
            time.sleep(5)
            print("模拟按下F5播放PPT")
            press_button(Key.f5)

            # 修复：守护循环逻辑正确
            while True:
                now_time = get_now_time()
                # 修复：已到结束时间 → 关闭WPS并退出
                if now_time > end_time:
                    print("晚自习结束，关闭WPS")
                    subprocess.run('''mshta vbscript:CreateObject("Wscript.Shell").popup("晚自习结束了",10,"WPS",64)(window.close)''')
                    subprocess.run("taskkill /f /im wpp.exe", shell=True)
                    break
                
                # 修复：WPS被关闭 → 自动重启+自动播放
                if not check_process_exist("wpp.exe"):
                    print("检测到WPS关闭，自动重启")
                    subprocess.run('''mshta vbscript:CreateObject("Wscript.Shell").popup("WPS演示被关闭，即将启动",10,"WPS",64)(window.close)''')
                    subprocess.Popen([wpp_path, ppt_path])
                    time.sleep(5)
                    press_button(Key.f5)
                
                # 修复：统一休眠，降低CPU占用
                time.sleep(5)
            # 修复：执行完成退出线程
            break
        
        # 未到时间，每秒检测一次
        time.sleep(1)

def get_wpp_path():
    reg_paths = [
        r"SOFTWARE\Kingsoft\Office\6.0\common",
        r"SOFTWARE\Wow6432Node\Kingsoft\Office\6.0\common"
    ]
    for reg_path in reg_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                install_path = winreg.QueryValueEx(key, "InstallPath")[0]
                wpp_path = os.path.join(install_path, "office6", "wpp.exe")
                if os.path.exists(wpp_path):
                    return wpp_path
        except:
            continue
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_paths[0]) as key:
            install_path = winreg.QueryValueEx(key, "InstallPath")[0]
            wpp_path = os.path.join(install_path, "office6", "wpp.exe")
            if os.path.exists(wpp_path):
                return wpp_path
    except:
        pass
    return None

if __name__ == '__main__':
    print("晚自习自动程序已启动")
    subprocess.Popen("taskkill /f /im HiteSmartScreenPro.exe")
    subprocess.Popen("taskkill /f /im HiteSmartScreenService.exe")
    subprocess.Popen("taskkill /f /im HiteSmartScreenSrv.exe")
    ppt_path = "D:\\静下心来专注学习.pptx"
    wpp_path = get_wpp_path()
    start_time = 65700
    end_time = 78600

    if not wpp_path:
        print("未读取到WPS路径，使用备用路径")
        wpp_path = r"C:\Program Files (x86)\Kingsoft\WPS Office\12.8.2.17149\office6\wpp.exe"
    print("WPS路径加载完成")

    # 修复：设置守护线程，主线程退出子线程自动关闭
    time_check_thread = threading.Thread(target=time_check, args=(wpp_path,ppt_path,start_time,end_time), daemon=True)
    # subprocess.run('''mshta vbscript:CreateObject("Wscript.Shell").popup("启动优化执行完毕",2,"WPS",64)(window.close)''')
    time.sleep(4)
    print("启动时间监控线程")
    time_check_thread.start()

    # 修复：删除join()，防止卡死；主线程保持运行
    while True:
        time.sleep(3600)