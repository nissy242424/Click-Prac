import yaml
import logging
from pynput import keyboard
import pyautogui
import mouse
import pywinauto
import ctypes
import sys

# 設定ファイル読み込み
with open('config.yaml') as f:
    config = yaml.safe_load(f)

# ログ設定
logging.basicConfig(
    filename=config['logging']['file_path'],
    level=config['logging']['level'],
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class MouseController:
    def __init__(self):
        self.trigger_key = config['key_bindings']['trigger_key']
        self.click_type = config['key_bindings']['click_type']
        self.libraries = config['libraries']['priority']
        self.current_library = None
        self.init_library()

    def init_library(self):
        """優先順位に従ってライブラリを初期化"""
        for lib in self.libraries:
            try:
                if lib == 'pynput':
                    self.current_library = 'pynput'
                    break
                elif lib == 'PyAutoGUI':
                    pyautogui.FAILSAFE = False
                    self.current_library = 'PyAutoGUI'
                    break
                elif lib == 'mouse':
                    self.current_library = 'mouse'
                    break
                elif lib == 'pywinauto':
                    self.current_library = 'pywinauto'
                    break
                elif lib == 'ctypes':
                    self.current_library = 'ctypes'
                    break
            except Exception as e:
                logging.warning(f'{lib} initialization failed: {str(e)}')
                continue

        if not self.current_library:
            logging.error('No available mouse control library')
            sys.exit(1)

    def perform_click(self):
        """マウスクリックを実行"""
        try:
            if self.current_library == 'pynput':
                # pynputはキー監視のみでクリック機能なし
                self.perform_click_with_pyautogui()
            elif self.current_library == 'PyAutoGUI':
                self.perform_click_with_pyautogui()
            elif self.current_library == 'mouse':
                mouse.click(button=self.click_type)
            elif self.current_library == 'pywinauto':
                pywinauto.mouse.click(button=self.click_type)
            elif self.current_library == 'ctypes':
                self.perform_click_with_ctypes()
            logging.info(f'Mouse click performed using {self.current_library}')
        except Exception as e:
            logging.error(f'Click failed: {str(e)}')
            self.init_library()  # フォールバック

    def perform_click_with_pyautogui(self):
        """PyAutoGUIを使用したクリック処理"""
        if self.click_type == 'left':
            pyautogui.click()
        elif self.click_type == 'right':
            pyautogui.rightClick()
        elif self.click_type == 'middle':
            pyautogui.middleClick()

    def perform_click_with_ctypes(self):
        """ctypesを使用したクリック処理"""
        # Windows APIを使用したクリック処理
        if self.click_type == 'left':
            ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # 左クリック押下
            ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # 左クリック解放
        elif self.click_type == 'right':
            ctypes.windll.user32.mouse_event(8, 0, 0, 0, 0)  # 右クリック押下
            ctypes.windll.user32.mouse_event(16, 0, 0, 0, 0)  # 右クリック解放

    def on_press(self, key):
        """キー押下時の処理"""
        try:
            key_char = key.char if hasattr(key, 'char') else str(key)
            logging.info(f'キー入力検出: {key_char}')

            if key_char == self.trigger_key:    
                logging.info(f'トリガーキー[{self.trigger_key}]が押されました。クリックを実行します。')
                self.perform_click()
            elif key == keyboard.Key.esc:  # ESCキーで終了
                logging.info('ESC key pressed. Exiting program...')
                self.stop()
                return False  # Listenerを停止
        except Exception as e:
            logging.error(f'Key press handling error: {str(e)}')

    def stop(self):
        """プログラム終了処理"""
        logging.info('Performing cleanup before exit')
        # 必要に応じてリソース解放処理を追加
        sys.exit(0)

    def start(self):
        """キーリスナーを開始"""
        with keyboard.Listener(on_press=self.on_press) as listener:
            logging.info('Mouse controller started')
            try:
                listener.join()
            except KeyboardInterrupt:
                self.stop()

if __name__ == '__main__':
    controller = MouseController()
    controller.start()
