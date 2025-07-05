from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from RecaptchaSolver import RecaptchaSolver
import time
from pywinauto import Desktop
import pyautogui
import tkinter as tk
from tkinter import messagebox, scrolledtext
import ctypes
import os
import sys
import threading
from queue import Queue

class RecaptchaBypassGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("web_click4 輔助工具")
        self.root.geometry("500x400")
        self.root.configure(bg="#f0f0f0")
        
        # 設定圖標
        try:
            # 主視窗圖示
            icon_path = "icon.ico"
            self.root.iconbitmap(icon_path)
            
            # 工作列圖示 - Windows 特定方法
            myappid = 'mycompany.recaptchabypass.1.0' # 任意字符串，但要唯一
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            
            # 另一種嘗試 - 使用 PhotoImage
            icon_img = tk.PhotoImage(file='icon.png')  # 需要 PNG 檔案
            self.root.tk.call('wm', 'iconphoto', self.root._w, icon_img)
        except Exception as e:
            print(f"設定圖示時出錯: {e}")
        
        # 主框架
        self.main_frame = tk.Frame(root, bg="#f0f0f0", padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 標題
        title_label = tk.Label(self.main_frame, 
                              text="Recaptcha 自動驗證工具", 
                              font=("Arial", 16, "bold"),
                              bg="#f0f0f0", fg="#333333")
        title_label.pack(pady=10)
        
        # 說明文字
        instruction = tk.Label(self.main_frame, 
                              text="請輸入Chrome偵錯端口，然後點擊開始執行",
                              font=("Arial", 10),
                              bg="#f0f0f0", fg="#555555")
        instruction.pack(pady=5)
        
        # 端口輸入區域
        input_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        input_frame.pack(fill=tk.X, pady=10)
        
        port_label = tk.Label(input_frame, text="Chrome 偵錯端口:",
                             font=("Arial", 10),
                             bg="#f0f0f0", fg="#333333")
        port_label.pack(side=tk.LEFT, padx=5)
        
        self.port_entry = tk.Entry(input_frame, font=("Arial", 10),
                                  width=10, bd=2, relief=tk.GROOVE)
        self.port_entry.pack(side=tk.LEFT, padx=5)
        
        # 按鈕區域
        button_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, pady=10)
        
        self.start_button = tk.Button(button_frame, 
                                    text="開始執行",
                                    command=self.start_monitoring,
                                    bg="#4CAF50", fg="white",
                                    font=("Arial", 10, "bold"),
                                    width=12, height=1,
                                    relief=tk.RAISED, bd=2)
        self.start_button.pack(side=tk.LEFT, padx=10, expand=True)
        
        self.stop_button = tk.Button(button_frame, 
                                   text="停止執行",
                                   command=self.stop_monitoring,
                                   bg="#f44336", fg="white",
                                   font=("Arial", 10, "bold"),
                                   width=12, height=1,
                                   relief=tk.RAISED, bd=2,
                                   state=tk.DISABLED)  # 一開始設為禁用
        self.stop_button.pack(side=tk.LEFT, padx=10, expand=True)
        
        # 狀態顯示區域 - 使用 ScrolledText
        status_frame = tk.LabelFrame(self.main_frame, text="執行狀態",
                                   font=("Arial", 10, "bold"),
                                   bg="#f0f0f0", fg="#333333",
                                   padx=10, pady=10)
        status_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.status_text = scrolledtext.ScrolledText(status_frame, 
                                                   wrap=tk.WORD,
                                                   height=10,
                                                   font=("Consolas", 9),
                                                   bg="#ffffff",
                                                   fg="#333333")
        self.status_text.pack(fill=tk.BOTH, expand=True)
        
        # 底部狀態列
        self.status_bar = tk.Label(self.main_frame, 
                                  text="準備就緒",
                                  bd=1, relief=tk.SUNKEN, anchor=tk.W,
                                  font=("Arial", 8),
                                  bg="#e0e0e0", fg="#555555")
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=5)
        
        # 初始化變數
        self.running = False
        self.driver = None
        self.recaptchaSolver = None
        
        # 在狀態文本中添加初始消息
        self.log_message("歡迎使用 Recaptcha 自動驗證工具")
        self.log_message("請輸入Chrome的偵錯端口，並點擊開始執行")
        self.log_message("例如: 9222, 或從DevTools輸出獲取")
        
        # 綁定按鈕懸停效果
        self.start_button.bind("<Enter>", lambda e: self.start_button.config(bg="#45a049"))
        self.start_button.bind("<Leave>", lambda e: self.start_button.config(bg="#4CAF50"))
        self.stop_button.bind("<Enter>", lambda e: self.stop_button.config(bg="#d32f2f") if self.stop_button["state"] == tk.NORMAL else None)
        self.stop_button.bind("<Leave>", lambda e: self.stop_button.config(bg="#f44336") if self.stop_button["state"] == tk.NORMAL else None)

        # 嘗試額外的方法設定工作列圖示
        try:
            if os.name == 'nt':
                # 確保使用絕對路徑
                if getattr(sys, 'frozen', False):
                    base_path = os.path.dirname(sys.executable)
                else:
                    base_path = os.path.dirname(os.path.abspath(__file__))
                
                icon_path = os.path.join(base_path, 'icon.ico')
                if os.path.exists(icon_path):
                    self.root.wm_iconbitmap(default=icon_path)
        except Exception as e:
            print(f"設定工作列圖示出錯 (在類中): {e}")

        # 初始化消息隊列
        self.log_queue = Queue()
        self.status_update_queue = Queue()

    def log_message(self, message):
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update()

    def create_driver_with_retry(self, debug_port, max_retries=3, retry_delay=2):
        for attempt in range(max_retries):
            try:
                chrome_options = Options()
                chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
                driver = webdriver.Chrome(options=chrome_options)
                self.queue_log_message(f"成功連接到 Chrome 實例 (Port: {debug_port})")
                return driver
            except Exception as e:
                self.queue_log_message(f"連接失敗 (嘗試 {attempt + 1}/{max_retries}): {str(e)}")
                if attempt < max_retries - 1:
                    self.queue_log_message(f"等待 {retry_delay} 秒後重試...")
                    time.sleep(retry_delay)
        raise Exception("無法連接到 Chrome 實例，已超過最大重試次數")

    def solve_captcha(self):
        try:
            t0 = time.time()
            self.recaptchaSolver.solveCaptcha()
            self.queue_log_message(f"驗證碼解決時間: {time.time()-t0:.2f} 秒")
        except Exception as e:
            self.queue_log_message(f"解決驗證碼時發生錯誤：{str(e)}")

    def check_for_dialog(self):
        """非阻塞版本的對話框檢測"""
        try:
            dialog = Desktop(backend="uia").window(title="警告")
            if dialog.exists():
                self.queue_log_message("檢測到警告對話框")
                dialog.set_focus()
                time.sleep(0.1)
                try:
                    ok_button = dialog.child_window(title="確定", control_type="Button")
                    if ok_button.exists():
                        ok_button.click()
                        self.queue_log_message("已點擊確定按鈕")
                    else:
                        pyautogui.press('enter')
                        self.queue_log_message("未找到確定按鈕，已按下 Enter 鍵")
                except:
                    pyautogui.press('enter')
                    self.queue_log_message("點擊按鈕失敗，已按下 Enter 鍵")
                
                self.solve_captcha()
                return True
        except Exception as e:
            self.queue_log_message(f"檢測對話框時發生錯誤：{str(e)}")
        
        return False

    def check_for_dialog2(self):
        """非阻塞版本的對話框檢測"""
        try:
            dialog = Desktop(backend="uia").window(title="請手動驗證不是機器人")
            if dialog.exists():
                self.queue_log_message("檢測到驗證對話框")
                dialog.set_focus()
                time.sleep(0.1)
                try:
                    ok_button = dialog.child_window(title="確定", control_type="Button")
                    if ok_button.exists():
                        ok_button.click()
                        self.queue_log_message("已點擊確定按鈕")
                    else:
                        pyautogui.press('enter')
                        self.queue_log_message("未找到確定按鈕，已按下 Enter 鍵")
                except:
                    pyautogui.press('enter')
                    self.queue_log_message("點擊按鈕失敗，已按下 Enter 鍵")
                return True
        except Exception as e:
            self.queue_log_message(f"檢測對話框時發生錯誤：{str(e)}")
        
        return False

    def monitoring_loop(self):
        # 創建一個單獨的線程運行監控循環
        self.monitor_thread = threading.Thread(target=self._monitor_thread_func)
        self.monitor_thread.daemon = True  # 設置為守護線程，主線程結束時會自動終止
        self.monitor_thread.start()
        
        # 定期更新 GUI (每 100ms)
        self.check_queue()

    def _monitor_thread_func(self):
        """在單獨線程中運行的監控函數"""
        while self.running:
            try:
                # 檢查 driver 是否還有效
                try:
                    self.driver.current_url
                except:
                    self.queue_log_message("檢測到連線中斷，嘗試重新連接...")
                    self.driver = self.create_driver_with_retry(self.port_entry.get())
                    self.recaptchaSolver = RecaptchaSolver(self.driver, log_callback=self.queue_log_message)
                
                if self.check_for_dialog():  # 修改對話框檢測邏輯
                    time.sleep(1)
                    self.check_for_dialog2()
                
                time.sleep(1)
                # 更新狀態欄需要在主線程中執行
                self.status_update_queue.put("監控中...")
                
            except Exception as e:
                self.queue_log_message(f"執行過程中發生錯誤：{str(e)}")
                continue

    def queue_log_message(self, message):
        """將日誌信息放入隊列，而不是直接更新 UI"""
        self.log_queue.put(message)

    def check_queue(self):
        """檢查並處理日誌和狀態更新隊列"""
        # 處理所有排隊的日誌消息
        try:
            while not self.log_queue.empty():
                message = self.log_queue.get_nowait()
                self.log_message(message)
        except:
            pass
        
        # 處理狀態欄更新
        try:
            while not self.status_update_queue.empty():
                status = self.status_update_queue.get_nowait()
                self.status_bar.config(text=status)
        except:
            pass
        
        # 如果仍在運行，繼續檢查隊列
        if self.running:
            self.root.after(100, self.check_queue)

    def start_monitoring(self):
        if not self.port_entry.get():
            messagebox.showerror("錯誤", "請輸入Chrome偵錯端口")
            return
        
        try:
            self.running = True
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.port_entry.config(state=tk.DISABLED)
            
            self.status_bar.config(text="正在啟動...")
            self.log_message("開始執行自動驗證...")
            
            # 在單獨的線程中初始化 driver
            threading.Thread(target=self._init_driver).start()
            
        except Exception as e:
            self.log_message(f"啟動失敗：{str(e)}")
            self.stop_monitoring()

    def _init_driver(self):
        """在單獨線程中初始化 driver"""
        try:
            self.driver = self.create_driver_with_retry(self.port_entry.get())
            self.recaptchaSolver = RecaptchaSolver(self.driver, log_callback=self.queue_log_message)
            
            # 啟動監控迴圈
            self.monitoring_loop()
        except Exception as e:
            self.queue_log_message(f"初始化失敗：{str(e)}")
            # 通知主線程停止監控
            self.root.after(0, self.stop_monitoring)

    def stop_monitoring(self):
        self.running = False
        
        # 在 GUI 線程中更新 UI
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.port_entry.config(state=tk.NORMAL)
        
        # 在單獨線程中關閉 driver
        if self.driver:
            threading.Thread(target=self._close_driver).start()
        
        self.log_message("已停止執行")
        self.status_bar.config(text="已停止")

    def _close_driver(self):
        """在單獨線程中關閉 driver"""
        try:
            self.driver.quit()
        except:
            pass
        self.driver = None

def main():
    # 確保在程式初始化前設定 AppUserModelID
    if os.name == 'nt':  # 確認是 Windows 系統
        # 獲取 exe 的絕對路徑，當打包成 exe 時非常重要
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        # 設定唯一的 AppUserModelID
        myappid = u'ecco.recaptchabypass.1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        
        # 確保圖示檔案存在
        icon_path = os.path.join(application_path, 'icon.ico')
        if not os.path.exists(icon_path):
            print(f"圖示檔案不存在: {icon_path}")
    
    # 創建 tkinter 視窗
    root = tk.Tk()
    
    # 嘗試直接設定 taskbar 圖示
    try:
        if os.name == 'nt':
            root.wm_iconbitmap(icon_path)
            # Windows 10/11 的備用方法
            root.after(10, lambda: ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid))
    except Exception as e:
        print(f"設定工作列圖示出錯: {e}")
    
    app = RecaptchaBypassGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
