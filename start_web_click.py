import subprocess
import os
import sys
import time
import psutil
import requests
from pywinauto.application import Application
from pywinauto.keyboard import send_keys


def find_listening_ports():
    """找出所有正在監聽的本地端口"""
    listening_ports = []
    for conn in psutil.net_connections(kind="inet"):
        if conn.status == "LISTEN" and conn.laddr.ip == "127.0.0.1":
            listening_ports.append(conn.laddr.port)
    return listening_ports


def find_process_by_port(port):
    """找出哪個進程在使用指定的端口"""
    for conn in psutil.net_connections(kind="inet"):
        if conn.laddr.port == port and conn.status == "LISTEN":
            try:
                return psutil.Process(conn.pid)
            except:
                return None
    return None


def is_browser_process(process):
    """判斷一個進程是否是瀏覽器進程（而不是ChromeDriver）"""
    if not process:
        return False

    name = process.name().lower()
    if "chromedriver" in name:
        return False

    if name in ["chrome.exe", "chromium.exe", "msedge.exe", "firefox.exe"]:
        return True

    return False


def test_devtools_connection(port):
    """測試是否可以連接到指定端口的DevTools"""
    try:
        url = f"http://127.0.0.1:{port}/json/version"
        response = requests.get(url, timeout=2)
        if response.status_code == 200:
            data = response.json()
            if "Browser" in data or "Protocol-Version" in data:
                return True
    except:
        pass
    return False


def wait_for_ui_loaded(exe_path, max_wait=600):
    """等待UI加載完成"""
    print("等待UI加載...")
    start_time = time.time()

    while time.time() - start_time < max_wait:
        try:
            app = Application(backend="win32").connect(path=exe_path)
            for window in app.windows():
                if window.is_visible():
                    print(f"找到主窗口: {window.window_text()}")
                    return app, window
        except:
            pass

        time.sleep(0.5)

    print("警告: 等待UI加載超時")
    return None, None


def monitor_for_devtools_port(old_ports, max_wait=600):
    """監控並等待DevTools端口出現"""
    print("開始監控新打開的端口...")
    start_time = time.time()
    new_ports = []
    devtools_port = None

    while time.time() - start_time < max_wait:
        # 檢查新端口
        current_ports = find_listening_ports()
        for port in current_ports:
            if port not in old_ports and port not in new_ports:
                print(f"發現新端口: {port}")
                new_ports.append(port)

                # 檢查進程類型
                process = find_process_by_port(port)
                if process:
                    process_name = process.name()
                    print(f"端口 {port} 由進程 {process.pid} ({process_name}) 使用")

                    if "chromedriver" in process_name.lower():
                        print(f"【識別】端口 {port} 是ChromeDriver端口")
                    elif is_browser_process(process):
                        print(f"【識別】端口 {port} 是瀏覽器進程端口")
                        # 瀏覽器進程的端口更可能是DevTools端口
                        devtools_port = port

                # 測試DevTools連接
                if test_devtools_connection(port):
                    print(f"【確認】端口 {port} 是DevTools端口 (通過API測試)")
                    devtools_port = port
                    return new_ports, devtools_port

        # 如果已經找到了兩個新端口，並且其中一個是瀏覽器進程的端口，可能就是我們要找的
        if len(new_ports) >= 2 and devtools_port:
            return new_ports, devtools_port

        time.sleep(1)

    # 如果超時但找到了多個端口，嘗試推測哪個是DevTools端口
    if len(new_ports) >= 2:
        # 假設最後一個新端口是DevTools端口
        devtools_port = new_ports[-1]

    return new_ports, devtools_port


def get_application_path():
    """獲取應用程序的實際路徑"""
    if getattr(sys, "frozen", False):
        # 如果是打包後的 exe
        return os.path.dirname(sys.executable)
    else:
        # 如果是 Python 腳本
        return os.path.dirname(os.path.abspath(__file__))


def start_web_click_and_press_start():
    # 使用新的路徑獲取方法
    application_path = get_application_path()

    # 構建web_click4.exe的完整路徑
    exe_path = os.path.join(application_path, "web_click4.exe")

    # 檢查文件是否存在
    if not os.path.exists(exe_path):
        sys.exit(1)

    # 在啟動程序前獲取當前監聽的端口
    find_listening_ports()

    # 啟動web_click4.exe
    print(f"正在啟動 {exe_path}...")
    subprocess.Popen(exe_path)

    # 等待UI加載
    _app, main_window = wait_for_ui_loaded(exe_path)

    if main_window:
        try:
            # 設置焦點到主窗口
            main_window.set_focus()
            time.sleep(1)  # 等待窗口獲取焦點

            # 按下Ctrl+Tab五次導航到"開始"按鈕
            for _i in range(5):
                send_keys("^{TAB}")
                time.sleep(0.5)

            # 按下空白鍵來點擊按鈕
            send_keys("{SPACE}")

            # 監控新打開的端口直到找到DevTools端口
            ports_before_click = find_listening_ports()
            new_ports, devtools_port = monitor_for_devtools_port(ports_before_click)

            if new_ports:
                if devtools_port:
                    return devtools_port
                else:
                    print("未能確定哪個是DevTools端口")
            else:
                print("點擊後沒有發現新端口")

        except Exception as e:
            print(f"操作時出錯: {e}")
            print("請手動點擊'開始'按鈕")
    else:
        print("警告: 無法找到主窗口")

    print("程序執行完畢")


if __name__ == "__main__":
    start_web_click_and_press_start()
