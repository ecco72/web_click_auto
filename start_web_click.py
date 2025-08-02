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


def find_window_by_process_name(process_name, max_wait=30, log_callback=None):
    """通過進程名稱查找窗口，更可靠的方法"""
    if log_callback:
        log_callback(f"正在查找進程 {process_name} 的窗口...")
    start_time = time.time()
    
    try:
        import win32gui
        import win32process
        
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process = psutil.Process(pid)
                    if process.name().lower() == process_name.lower():
                        window_text = win32gui.GetWindowText(hwnd)
                        windows.append((hwnd, window_text, pid))
                except:
                    pass
            return True
        
        while time.time() - start_time < max_wait:
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            if windows:
                for hwnd, window_text, pid in windows:
                    if log_callback:
                        log_callback(f"找到窗口: {window_text} (PID: {pid})")
                    return hwnd, window_text
            
            time.sleep(0.5)
    except ImportError:
        if log_callback:
            log_callback("win32gui 不可用，使用備用方法")
    
    if log_callback:
        log_callback(f"未找到 {process_name} 的窗口")
    return None, None


def wait_for_ui_loaded(exe_path, max_wait=600, log_callback=None):
    """改進的UI加載等待函數"""
    if log_callback:
        log_callback("等待UI加載...")
    start_time = time.time()
    exe_name = os.path.basename(exe_path)
    
    # 方法1: 通過進程名稱查找窗口
    hwnd, window_text = find_window_by_process_name(exe_name, max_wait=30, log_callback=log_callback)
    if hwnd:
        try:
            app = Application(backend="win32").connect(handle=hwnd)
            window = app.window(handle=hwnd)
            if log_callback:
                log_callback(f"成功連接到窗口: {window_text}")
            return app, window
        except Exception as e:
            if log_callback:
                log_callback(f"連接窗口失敗: {e}")
    
    # 方法2: 原始方法作為備用
    if log_callback:
        log_callback("嘗試使用備用方法查找窗口...")
    while time.time() - start_time < max_wait:
        try:
            app = Application(backend="win32").connect(path=exe_path)
            for window in app.windows():
                if window.is_visible():
                    if log_callback:
                        log_callback(f"找到主窗口: {window.window_text()}")
                    return app, window
        except:
            pass

        # 方法3: 嘗試通過標題查找
        try:
            app = Application(backend="win32").connect(title_re=".*web.*click.*")
            for window in app.windows():
                if window.is_visible():
                    if log_callback:
                        log_callback(f"通過標題找到窗口: {window.window_text()}")
                    return app, window
        except:
            pass

        time.sleep(0.5)

    if log_callback:
        log_callback("警告: 等待UI加載超時")
    return None, None


def monitor_for_devtools_port(old_ports, max_wait=600, log_callback=None):
    """監控並等待DevTools端口出現"""
    if log_callback:
        log_callback("開始監控新打開的端口...")
    start_time = time.time()
    new_ports = []
    devtools_port = None

    while time.time() - start_time < max_wait:
        # 檢查新端口
        current_ports = find_listening_ports()
        for port in current_ports:
            if port not in old_ports and port not in new_ports:
                if log_callback:
                    log_callback(f"發現新端口: {port}")
                new_ports.append(port)

                # 檢查進程類型
                process = find_process_by_port(port)
                if process:
                    process_name = process.name()
                    if log_callback:
                        log_callback(f"端口 {port} 由進程 {process.pid} ({process_name}) 使用")

                    if "chromedriver" in process_name.lower():
                        if log_callback:
                            log_callback(f"【識別】端口 {port} 是ChromeDriver端口")
                    elif is_browser_process(process):
                        if log_callback:
                            log_callback(f"【識別】端口 {port} 是瀏覽器進程端口")
                        # 瀏覽器進程的端口更可能是DevTools端口
                        devtools_port = port

                # 測試DevTools連接
                if test_devtools_connection(port):
                    if log_callback:
                        log_callback(f"【確認】端口 {port} 是DevTools端口 (通過API測試)")
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


def enable_lock_screen_execution(log_callback=None):
    """啟用鎖定畫面執行的權限設置"""
    try:
        # 設置進程為高優先級
        try:
            import win32api
            import win32process
            handle = win32api.GetCurrentProcess()
            win32process.SetPriorityClass(handle, win32process.HIGH_PRIORITY_CLASS)
            if log_callback:
                log_callback("已設置高優先級")
        except ImportError:
            if log_callback:
                log_callback("win32api 不可用，跳過優先級設置")
        
        # 嘗試獲取更多權限
        try:
            import win32security
            token = win32security.OpenProcessToken(
                win32api.GetCurrentProcess(),
                win32security.TOKEN_ADJUST_PRIVILEGES | win32security.TOKEN_QUERY
            )
            privilege = win32security.LookupPrivilegeValue(None, win32security.SE_DEBUG_NAME)
            win32security.AdjustTokenPrivileges(
                token, 0, [(privilege, win32security.SE_PRIVILEGE_ENABLED)]
            )
            if log_callback:
                log_callback("已啟用調試權限")
        except (ImportError, Exception) as e:
            if log_callback:
                log_callback(f"啟用高級權限失敗: {e}")
        
        return True
    except Exception as e:
        if log_callback:
            log_callback(f"設置執行權限失敗: {e}")
        return False


def send_keys_with_retry(keys, retries=3, log_callback=None):
    """帶重試的按鍵發送"""
    for attempt in range(retries):
        try:
            send_keys(keys)
            return True
        except Exception as e:
            if log_callback:
                log_callback(f"發送按鍵失敗 (嘗試 {attempt + 1}/{retries}): {e}")
            time.sleep(0.5)
    return False


def click_start_button_safe(main_window, log_callback=None):
    """安全的點擊開始按鈕，支援鎖定畫面"""
    try:
        if log_callback:
            log_callback("正在搜尋開始按鈕...")
        
        
        # 方法1: 直接座標點擊（適用於所有情況）
        try:
            rect = main_window.rectangle()
            # 計算可能的按鈕位置（窗口中央偏下）
            click_x = rect.left + (rect.width() // 2)
            click_y = rect.top + int(rect.height() * 0.7)  # 70%的位置
            
            if log_callback:
                log_callback(f"嘗試點擊座標 ({click_x}, {click_y})")
            
            # 使用相對座標點擊
            main_window.click_input(coords=(rect.width() // 2, int(rect.height() * 0.7)))
            if log_callback:
                log_callback("成功點擊座標位置")
            return True
        except Exception as e:
            if log_callback:
                log_callback(f"座標點擊失敗: {e}")
        
        # 方法2: 嘗試通過按鈕文字查找
            try:
                start_button = main_window.child_window(title="開始", control_type="Button")
                if start_button.exists():
                    start_button.click()
                    if log_callback:
                        log_callback("通過文字'開始'找到並點擊按鈕")
                    return True
            except Exception as e:
                if log_callback:
                    log_callback(f"通過'開始'文字查找失敗: {e}")

        return False
        
    except Exception as e:
        if log_callback:
            log_callback(f"點擊開始按鈕時發生錯誤: {e}")
        return False


def click_start_button(main_window, log_callback=None):
    """嘗試直接點擊開始按鈕（保留原函數名稱以兼容）"""
    return click_start_button_safe(main_window, log_callback)


def activate_window_force(hwnd, log_callback=None):
    """強制激活窗口，即使在鎖定畫面"""
    try:
        import win32gui
        import win32con
        
        # 嘗試多種方法激活窗口
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.BringWindowToTop(hwnd)
        
        # 使用更強制的方法
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 
                             win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, 
                             win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        
        return True
    except (ImportError, Exception) as e:
        if log_callback:
            log_callback(f"激活窗口失敗: {e}")
        return False


def start_web_click_and_press_start(log_callback=None):
    """改進的啟動函數，支持鎖定畫面執行"""
    # 啟用鎖定畫面執行權限
    enable_lock_screen_execution(log_callback=log_callback)
    
    # 使用新的路徑獲取方法
    application_path = get_application_path()

    # 構建web_click5.exe的完整路徑
    exe_path = os.path.join(application_path, "web_click5.exe")

    # 檢查文件是否存在
    if not os.path.exists(exe_path):
        if log_callback:
            log_callback(f"找不到文件: {exe_path}")
        sys.exit(1)

    # 在啟動程序前獲取當前監聽的端口
    ports_before = find_listening_ports()
    if log_callback:
        log_callback(f"啟動前的端口: {ports_before}")

    # 啟動web_click5.exe
    if log_callback:
        log_callback(f"正在啟動 {exe_path}...")
    process = subprocess.Popen(exe_path)
    
    # 等待程序啟動
    time.sleep(2)

    # 等待UI加載
    _app, main_window = wait_for_ui_loaded(exe_path, log_callback=log_callback)

    if main_window:
        try:
            # 嘗試激活窗口（在鎖定畫面下可能失敗，但不影響後續操作）
            try:
                hwnd = main_window.handle
                if activate_window_force(hwnd, log_callback=log_callback):
                    if log_callback:
                        log_callback("窗口已激活")
                
                # 設置焦點到主窗口
                main_window.set_focus()
                time.sleep(1)  # 等待窗口獲取焦點
            except Exception as e:
                if log_callback:
                    log_callback(f"窗口激活失敗，但繼續執行: {e}")

            if log_callback:
                log_callback("正在尋找並點擊開始按鈕...")
            
            # 嘗試直接點擊開始按鈕
            button_clicked = False
            if click_start_button(main_window, log_callback=log_callback):
                if log_callback:
                    log_callback("成功點擊開始按鈕")
                button_clicked = True
            else:
                if log_callback:
                    log_callback("座標點擊失敗，可能是鎖定畫面或其他原因")
                    log_callback("建議: 手動解鎖電腦或設定永不鎖定螢幕")
                    log_callback("程式將繼續監控端口變化...")
            
            # 等待一下讓按鈕點擊生效
            if button_clicked:
                time.sleep(2)
                if log_callback:
                    log_callback("等待程式響應按鈕點擊...")

            # 監控新打開的端口直到找到DevTools端口
            new_ports, devtools_port = monitor_for_devtools_port(ports_before, log_callback=log_callback)

            if new_ports:
                if log_callback:
                    log_callback(f"發現新端口: {new_ports}")
                if devtools_port:
                    if log_callback:
                        log_callback(f"確定DevTools端口: {devtools_port}")
                    return devtools_port
                else:
                    if log_callback:
                        log_callback("未能確定哪個是DevTools端口")
                    # 返回最後一個端口作為猜測
                    return new_ports[-1] if new_ports else None
            else:
                if log_callback:
                    log_callback("點擊後沒有發現新端口")

        except Exception as e:
            if log_callback:
                log_callback(f"操作時出錯: {e}")
                log_callback("請手動點擊'開始'按鈕")
    else:
        if log_callback:
            log_callback("警告: 無法找到主窗口，嘗試手動操作")
        # 即使找不到窗口，也監控端口變化
        time.sleep(5)  # 給用戶時間手動點擊
        new_ports, devtools_port = monitor_for_devtools_port(ports_before, max_wait=60, log_callback=log_callback)
        if devtools_port:
            return devtools_port

    if log_callback:
        log_callback("程序執行完畢")
    return None


if __name__ == "__main__":
    start_web_click_and_press_start()
