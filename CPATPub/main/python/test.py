import shutil
import os

# フォルダのパスとファイル名を指定
folder_path = "C:\\Users\\PowerSystemLab\\Documents\\FromDesktop\\01_研究資料\\05_実行ファイル\\CPAT\\02_MyModel\\CPATFreeR05V1.0\\CPATFree_R05_Ver1.0\\model\\EAST10\\main"
file_name = "EAST10peak_AllThermal_AllDemand_auto.pop"

# 元ファイルのパス
source_path = os.path.join(folder_path, file_name)

# コピー先のパス（コピー後のファイル名を指定）
destination_path = os.path.join(folder_path, "EAST10peak_AllThermal_AllDemand_auto_copy.pop")

# ファイルをコピー
shutil.copy2(source_path, destination_path)

print(f"ファイル {file_name} を {destination_path} にコピーしました。")

################################################################################################################################

from pynput import mouse, keyboard
import time
import json
import os

def replay_actions(file_name):
    with open(file_name, 'r') as f:
        actions = json.load(f)

    start_time = actions[0][-1]

    mouse_controller = mouse.Controller()
    keyboard_controller = keyboard.Controller()

    for action in actions:
        action_type = action[0]
        if action_type == 'move':
            _, x, y, timestamp = action
            mouse_controller.position = (x, y)
        elif action_type == 'click':
            _, x, y, button, pressed, timestamp = action
            mouse_controller.position = (x, y)
            button = mouse.Button[button.split('.')[1]]
            if pressed:
                mouse_controller.press(button)
            else:
                mouse_controller.release(button)
        elif action_type == 'scroll':
            _, x, y, dx, dy, timestamp = action
            mouse_controller.scroll(dx, dy)
        elif action_type == 'key':
            _, key, event_type, timestamp = action
            try:
                if key.startswith('Key.'):
                    key = getattr(keyboard.Key, key.split('.')[1])
                if event_type == 'press':
                    keyboard_controller.press(key)
                else:
                    keyboard_controller.release(key)
            except AttributeError:
                print(f"未対応のキー: {key}")
            except ValueError as e:
                print(f"キーの解決中にエラーが発生しました: {e}")

        time.sleep(max(0, timestamp - start_time))
        start_time = timestamp

def replay_all_actions(file_names):
    for file_name in file_names:
        print(f"再生中: {file_name}")
        replay_actions(file_name)
        print(f"再生完了: {file_name}")

# 再生するファイル名のリストを作成
# action_files = ["OpenCPAT.json"]  # ここに再生したいファイル名を追加
action_files = ["OpenCPAT.json"]  # ここに再生したいファイル名を追加,
# action_files = ["MakeFile.json"]  # ここに再生したいファイル名を追加,

# ファイルが存在するか確認
action_files = [file for file in action_files if os.path.exists(file)]

# 10回繰り返して実行
for _ in range(1):
    replay_all_actions(action_files)

import pygetwindow as gw
import pyautogui

# 現在のアクティブウィンドウを取得
window = gw.getActiveWindow()

# ウィンドウを最大化
if window is not None:
    window.maximize()
else:
    print("アクティブなウィンドウが見つかりませんでした。")

################################################################################################################################

import subprocess
import win32com.client
import pyperclip
import math

# Excelアプリケーションを起動
excel = win32com.client.Dispatch("Excel.Application")

# Excelファイルのパスを指定 (.xlsm拡張子を使用)
file_path = "C:\\Users\\PowerSystemLab\\Documents\\FromDesktop\\01_研究資料\\05_実行ファイル\\CPAT\\02_MyModel\\CPATFreeR05V1.0\\CPATFree_R05_Ver1.0\\model\\EAST10\\main\\condition.xlsm"

# Excelファイルを開く
workbook = excel.Workbooks.Open(file_path)

# アクティブなシートを取得
sheet = workbook.ActiveSheet

# セルB3の値を取得
b3_value = sheet.Range("O2").Value
b3_value = math.floor(b3_value)

# 取得した値を表示
print(f"セルO2の値: {b3_value}")

# ファイルを保存して閉じる
workbook.Save()
workbook.Close()

# 取得した値をクリップボードにコピー
pyperclip.copy(str(b3_value))

print(f"セルB3の値 '{b3_value}' をクリップボードにコピーしました。")

# Excelアプリケーションを終了
excel.Quit()

################################################################################################################################

# 再生するファイル名のリストを作成
# action_files = ["OpenCPAT.json"]  # ここに再生したいファイル名を追加
action_files = ["ExecuteCPAT.json"]  # ここに再生したいファイル名を追加,
# action_files = ["MakeFile.json"]  # ここに再生したいファイル名を追加,

# ファイルが存在するか確認
action_files = [file for file in action_files if os.path.exists(file)]

# 10回繰り返して実行
for _ in range(1):
    replay_all_actions(action_files)

################################################################################################################################

import pygetwindow as gw
import pyautogui
import time

# 現在のアクティブウィンドウを取得
window = gw.getActiveWindow()

if window is not None:
    # ウィンドウタイトルに"Excel"が含まれているか確認
    if 'Excel' in window.title:
        # Excelのウィンドウを保存して終了する手順
        pyautogui.hotkey('ctrl', 's')  # 保存
        time.sleep(1)  # 保存のための短い待ち時間
        window.close()  # ウィンドウを閉じる
    else:
        # 他のウィンドウをそのまま終了
        window.close()
else:
    print("アクティブなウィンドウが見つかりませんでした。")

################################################################################################################################

# 再生するファイル名のリストを作成
# action_files = ["OpenCPAT.json"]  # ここに再生したいファイル名を追加
action_files = ["CloseCPAT.json"]  # ここに再生したいファイル名を追加,
# action_files = ["MakeFile.json"]  # ここに再生したいファイル名を追加,

# ファイルが存在するか確認
action_files = [file for file in action_files if os.path.exists(file)]

# 10回繰り返して実行
for _ in range(1):
    replay_all_actions(action_files)

################################################################################################################################

import os

# 削除したいファイルのパスを指定
file_path = 'C:\\Users\\PowerSystemLab\\Documents\\FromDesktop\\01_研究資料\\05_実行ファイル\\CPAT\\02_MyModel\\CPATFreeR05V1.0\\CPATFree_R05_Ver1.0\\model\\EAST10\\main\\EAST10peak_AllThermal_AllDemand_auto_copy.pop'

# ファイルの存在を確認し、削除
if os.path.exists(file_path):
    os.remove(file_path)
    print(f"{file_path} を削除しました。")
else:
    print(f"{file_path} が見つかりません。")

import os
import shutil

# 移動および削除したいファイルのパス
file_path = 'C:\\Users\\PowerSystemLab\\Documents\\FromDesktop\\01_研究資料\\05_実行ファイル\\CPAT\\02_MyModel\\CPATFreeR05V1.0\\CPATFree_R05_Ver1.0\\Temp\\EAST10peak_AllThermal_AllDemand_auto_copy'
# 一時ディレクトリのパス
temp_dir = 'C:\\TempDelete'

# 一時ディレクトリが存在しない場合、作成
if not os.path.exists(temp_dir):
    os.makedirs(temp_dir)

# ファイルを一時ディレクトリに移動
temp_path = os.path.join(temp_dir, os.path.basename(file_path))
try:
    shutil.move(file_path, temp_path)
    print(f"{file_path} を {temp_path} に移動しました。")
    # 一時ディレクトリからファイルを削除
    os.remove(temp_path)
    print(f"{temp_path} を削除しました。")
except PermissionError as e:
    print(f"PermissionError: {e}")
except Exception as e:
    print(f"他のエラーが発生しました: {e}")
