from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import os
import urllib.request
import random
import pydub
import speech_recognition
import time
from typing import Optional
import subprocess
import sys
import io


class RecaptchaSolver:
    """A class to solve reCAPTCHA challenges using audio recognition."""

    # Constants
    TEMP_DIR = os.getenv("TEMP") if os.name == "nt" else "/tmp"
    TIMEOUT_STANDARD = 7
    TIMEOUT_SHORT = 1
    TIMEOUT_DETECTION = 0.05

    def __init__(self, driver: webdriver.Chrome, log_callback=None) -> None:
        """Initialize the solver with a Selenium Chrome WebDriver.

        Args:
            driver: Selenium Chrome WebDriver instance
            log_callback: Callback function for logging messages
        """
        self.driver = driver
        self.wait = WebDriverWait(self.driver, 20)
        self.log_callback = log_callback

    def solveCaptcha(self) -> None:
        """Attempt to solve the reCAPTCHA challenge.

        Raises:
            Exception: If captcha solving fails or bot is detected
        """

        try:
            # Handle main reCAPTCHA iframe
            iframe = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']")
                )
            )
            self.driver.switch_to.frame(iframe)
            time.sleep(2)  # 給予更多時間讓元素載入

            # Click the checkbox
            checkbox = self.wait.until(
                EC.element_to_be_clickable((By.CLASS_NAME, "rc-anchor-content"))
            )
            checkbox.click()

            # Check if solved by just clicking
            if self.is_solved():
                return

            # Handle audio challenge
            # 切換回主頁面
            self.driver.switch_to.default_content()
            # 等待並切換到新的 reCAPTCHA iframe
            iframe = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "iframe[src*='recaptcha/api2/bframe']")
                )
            )
            self.driver.switch_to.frame(iframe)
            time.sleep(2)  # 給予更多時間讓元素載入

            # Click the audio button
            audio_button = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "button#recaptcha-audio-button.rc-button-audio")
                )
            )

            # 確保按鈕可以點擊
            if not audio_button.is_displayed() or not audio_button.is_enabled():
                time.sleep(2)  # 給予額外的等待時間

            # 使用 JavaScript 點擊（如果一般點擊失敗的話）
            try:
                audio_button.click()
            except Exception as e:
                self.driver.execute_script("arguments[0].click();", audio_button)
                if self.log_callback:
                    self.log_callback(f"使用 JavaScript 點擊按鈕: {str(e)}")

            time.sleep(0.3)

            if self.is_detected():
                raise Exception("Captcha detected bot behavior")

            # Download and process audio
            audio_source = self.wait.until(
                EC.presence_of_element_located((By.ID, "audio-source"))
            )
            src = audio_source.get_attribute("src")

            time.sleep(2)  # 給音頻加載更多時間

            text_response = self._process_audio_challenge(
                src, max_retries=5
            )  # 增加重試次數

            # 在輸入答案前後增加延遲
            time.sleep(1)  # 輸入前等待
            self.driver.find_element(By.ID, "audio-response").send_keys(
                text_response.lower()
            )
            time.sleep(1)  # 輸入後等待
            self.driver.find_element(By.ID, "recaptcha-verify-button").click()
            time.sleep(1)  # 驗證後等待更長時間

            # 增加驗證結果檢查的重試
            for _ in range(3):  # 最多重試3次
                if self.is_solved():
                    return
                time.sleep(1)

        except Exception:
            self.driver.switch_to.default_content()
            raise

        finally:
            self.driver.switch_to.default_content()

    def _process_audio_challenge(self, audio_url: str, max_retries: int = 3) -> str:
        """Process the audio challenge and return the recognized text."""

        # 創建一個空的輸出重定向對象
        class NullIO(io.IOBase):
            def write(self, *args, **kwargs):
                pass

            def read(self, *args, **kwargs):
                return ""

            def close(self):
                pass

            def flush(self):
                pass

        # 保存原始的標準輸出和錯誤輸出
        old_stdout = sys.stdout
        old_stderr = sys.stderr

        # 在 Windows 上設置 subprocess 的 STARTUPINFO 以隱藏控制台窗口
        if os.name == "nt":
            # 修改環境變數以影響子程序
            os.environ["PYTHONUNBUFFERED"] = "1"

            # 設置 subprocess 啟動信息
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE

            # 猴子補丁 subprocess.Popen 以使用我們的 startupinfo
            original_popen = subprocess.Popen

            def new_popen(*args, **kwargs):
                if "startupinfo" not in kwargs:
                    kwargs["startupinfo"] = startupinfo
                if "creationflags" not in kwargs:
                    kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
                return original_popen(*args, **kwargs)

            subprocess.Popen = new_popen

        for attempt in range(max_retries):
            mp3_path = None
            wav_path = None
            try:
                # 重定向標準輸出和錯誤輸出，防止cmd顯示
                sys.stdout = NullIO()
                sys.stderr = NullIO()

                mp3_path = os.path.join(
                    self.TEMP_DIR, f"{random.randrange(1, 1000)}.mp3"
                )
                wav_path = os.path.join(
                    self.TEMP_DIR, f"{random.randrange(1, 1000)}.wav"
                )

                # 使用urlretrieve下載音頻文件
                urllib.request.urlretrieve(audio_url, mp3_path)

                # 使用pydub轉換音頻格式 - 可能會啟動子進程
                sound = pydub.AudioSegment.from_mp3(mp3_path)
                sound.export(wav_path, format="wav")

                # 使用speech_recognition識別文字 - 可能會啟動子進程
                recognizer = speech_recognition.Recognizer()
                with speech_recognition.AudioFile(wav_path) as source:
                    audio = recognizer.record(source)

                # 使用Google API進行識別
                text = recognizer.recognize_google(audio)

                # 恢復標準輸出和錯誤輸出
                sys.stdout = old_stdout
                sys.stderr = old_stderr

                # 還原 subprocess.Popen
                if os.name == "nt":
                    subprocess.Popen = original_popen

                return text

            except speech_recognition.UnknownValueError:
                # 恢復標準輸出和錯誤輸出
                sys.stdout = old_stdout
                sys.stderr = old_stderr

                # 還原 subprocess.Popen
                if os.name == "nt":
                    subprocess.Popen = original_popen

                if attempt == max_retries - 1:
                    raise
                time.sleep(2)  # 在重試之前等待

            except Exception as e:
                # 恢復標準輸出和錯誤輸出
                sys.stdout = old_stdout
                sys.stderr = old_stderr

                # 還原 subprocess.Popen
                if os.name == "nt":
                    subprocess.Popen = original_popen

                if self.log_callback:
                    self.log_callback(f"音頻處理錯誤: {str(e)}")
                raise

            finally:
                # 恢復標準輸出和錯誤輸出 (以防有異常)
                sys.stdout = old_stdout
                sys.stderr = old_stderr

                # 還原 subprocess.Popen
                if os.name == "nt":
                    subprocess.Popen = original_popen

                # 清理臨時文件
                for path in [mp3_path, wav_path]:
                    if path and os.path.exists(path):
                        try:
                            os.remove(path)
                        except OSError:
                            pass

    def is_solved(self) -> bool:
        """Check if the captcha has been solved successfully."""
        try:
            self.driver.switch_to.default_content()
            iframe = self.driver.find_element(
                By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"
            )
            self.driver.switch_to.frame(iframe)
            return (
                self.driver.find_element(
                    By.CLASS_NAME, "recaptcha-checkbox-checkmark"
                ).get_attribute("style")
                != ""
            )
        except Exception:
            return False
        finally:
            self.driver.switch_to.default_content()

    def is_detected(self) -> bool:
        """Check if the bot has been detected."""
        try:
            return "Try again later" in self.driver.page_source
        except Exception:
            return False

    def get_token(self) -> Optional[str]:
        """Get the reCAPTCHA token if available."""
        try:
            return self.driver.find_element(By.ID, "recaptcha-token").get_attribute(
                "value"
            )
        except Exception:
            return None
