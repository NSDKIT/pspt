# =============================================================================
# CPAT GUI 操作モジュール (gui_actions.py)
# =============================================================================
# このモジュールは、CPAT-GUI に対する各種 GUI 操作を
# 「キーボードショートカット優先 + 画像認識フォールバック」で実装します。
#
# 【旧方式との違い】
#   旧方式: JSON マクロの座標クリック（画面解像度に依存）
#   新方式: キーボードショートカット + 画像認識（環境に依存しにくい）
#
# 【使い方】
#   from gui_actions import open_cpat, open_mdl, run_simulation, close_cpat
#   open_cpat(cpat_exe_path)
#   open_mdl(mdl_file_path)
#   run_simulation()
#   close_cpat()
# =============================================================================

import os
import time
import subprocess
import pyautogui
import pygetwindow as gw

# pyautogui の安全装置（マウスを画面左上に移動すると緊急停止）
pyautogui.FAILSAFE = True
# 操作間のデフォルト待機時間 [秒]
pyautogui.PAUSE = 0.3

# このファイルが置かれているフォルダ（画像ファイルの基点）
_HERE = os.path.dirname(os.path.abspath(__file__))
# 画像ファイルを置くフォルダ
_IMG_DIR = os.path.join(_HERE, "gui_images")


# =============================================================================
# 内部ユーティリティ
# =============================================================================

def _wait_window(title_keyword: str, timeout: float = 30) -> bool:
    """指定キーワードを含むウィンドウが現れるまで待機する。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        wins = [w for w in gw.getAllWindows() if title_keyword in w.title]
        if wins:
            return True
        time.sleep(0.5)
    print(f"[警告] ウィンドウ '{title_keyword}' が {timeout}s 以内に見つかりませんでした。")
    return False


def _activate_window(title_keyword: str) -> bool:
    """指定キーワードを含むウィンドウをアクティブにする。"""
    wins = [w for w in gw.getAllWindows() if title_keyword in w.title]
    if wins:
        try:
            wins[0].activate()
            time.sleep(0.5)
            return True
        except Exception as e:
            print(f"[警告] ウィンドウのアクティブ化に失敗: {e}")
    return False


def _click_image_or_fallback(image_name: str, fallback_keys: list = None,
                               confidence: float = 0.8, timeout: float = 10) -> bool:
    """画像認識でクリックを試み、失敗した場合はキーボード操作にフォールバックする。

    Args:
        image_name: gui_images/ フォルダ内の画像ファイル名（例: "btn_run.png"）
        fallback_keys: 画像が見つからなかった場合に送信するキーリスト
                       例: [('alt', 's'), ('enter',)]
        confidence: 画像一致度のしきい値
        timeout: 画像を探す最大時間 [秒]

    Returns:
        True: 操作成功 / False: 失敗
    """
    image_path = os.path.join(_IMG_DIR, image_name)

    # 画像認識を試みる
    if os.path.exists(image_path):
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                location = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if location:
                    center = pyautogui.center(location)
                    pyautogui.click(center)
                    time.sleep(0.5)
                    print(f"[画像認識] クリック成功: {image_name}")
                    return True
            except pyautogui.ImageNotFoundException:
                pass
            time.sleep(0.5)
        print(f"[画像認識] 画像が見つかりませんでした: {image_name}")
    else:
        print(f"[画像認識] 画像ファイルが存在しません: {image_path}")
        print("  → gui_images/ フォルダにスクリーンショットを配置してください。")

    # フォールバック: キーボード操作
    if fallback_keys:
        print(f"[フォールバック] キーボード操作を実行します: {fallback_keys}")
        for keys in fallback_keys:
            if isinstance(keys, (list, tuple)):
                pyautogui.hotkey(*keys)
            else:
                pyautogui.press(keys)
            time.sleep(0.3)
        return True

    return False


# =============================================================================
# CPAT 操作関数
# =============================================================================

def open_cpat(cpat_exe_path: str, timeout: float = 30) -> bool:
    """CPAT.exe を起動し、メインウィンドウが表示されるまで待機する。

    Args:
        cpat_exe_path: CPAT.exe の絶対パス
        timeout: 起動待機の最大時間 [秒]

    Returns:
        True: 起動成功 / False: タイムアウト
    """
    if not os.path.exists(cpat_exe_path):
        print(f"[エラー] CPAT.exe が見つかりません: {cpat_exe_path}")
        return False

    print(f"CPAT を起動します: {cpat_exe_path}")
    subprocess.Popen([cpat_exe_path])
    result = _wait_window("CPAT", timeout=timeout)
    if result:
        print("CPAT の起動を確認しました。")
    return result


def open_mdl(mdl_file_path: str, timeout: float = 30) -> bool:
    """CPAT でモデルファイル (.pop) を開く。

    キーボードショートカット Ctrl+O でファイルを開くダイアログを呼び出し、
    pyautogui.typewrite でパスを入力します。

    Args:
        mdl_file_path: 開くモデルファイルの絶対パス
        timeout: ダイアログ待機の最大時間 [秒]

    Returns:
        True: 成功 / False: 失敗
    """
    if not _activate_window("CPAT"):
        print("[エラー] CPAT ウィンドウが見つかりません。")
        return False

    print(f"モデルを開きます: {mdl_file_path}")
    # Ctrl+O でファイルを開くダイアログを呼び出す
    pyautogui.hotkey('ctrl', 'o')
    time.sleep(1.5)

    # ファイルパスを入力して Enter
    pyautogui.typewrite(mdl_file_path, interval=0.03)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(2)
    return True


def run_simulation(timeout: float = 300) -> bool:
    """CPAT でシミュレーションを実行する。

    メニュー操作の代わりにキーボードショートカットを使用します。
    CPAT のショートカットキーに合わせて適宜変更してください。

    Args:
        timeout: シミュレーション完了待機の最大時間 [秒]

    Returns:
        True: 実行開始 / False: 失敗
    """
    if not _activate_window("CPAT"):
        print("[エラー] CPAT ウィンドウが見つかりません。")
        return False

    print("シミュレーションを実行します。")

    # 方法1: 画像認識で「実行」ボタンをクリック（gui_images/btn_run.png が必要）
    success = _click_image_or_fallback(
        image_name="btn_run.png",
        fallback_keys=[('f5',)],  # フォールバック: F5 キー（CPAT の実行ショートカット）
        timeout=10
    )

    if success:
        print(f"シミュレーション実行中... （最大 {timeout}s 待機）")
        time.sleep(3)  # 実行開始を待つ
    return success


def save_result_csv(timeout: float = 30) -> bool:
    """シミュレーション結果を CSV として保存する。

    Args:
        timeout: 保存完了待機の最大時間 [秒]

    Returns:
        True: 成功 / False: 失敗
    """
    if not _activate_window("CPAT"):
        print("[エラー] CPAT ウィンドウが見つかりません。")
        return False

    print("結果を CSV として保存します。")

    # 方法1: 画像認識で「CSV保存」ボタンをクリック（gui_images/btn_save_csv.png が必要）
    success = _click_image_or_fallback(
        image_name="btn_save_csv.png",
        fallback_keys=[('ctrl', 's')],  # フォールバック: Ctrl+S
        timeout=10
    )
    if success:
        time.sleep(2)
    return success


def close_cpat() -> bool:
    """CPAT を閉じる。

    Returns:
        True: 成功 / False: 失敗
    """
    if not _activate_window("CPAT"):
        print("[警告] CPAT ウィンドウが見つかりません（既に閉じている可能性があります）。")
        return True

    print("CPAT を閉じます。")

    # Alt+F4 でウィンドウを閉じる
    pyautogui.hotkey('alt', 'F4')
    time.sleep(1)

    # 「保存しますか？」ダイアログが出た場合は「いいえ」を選択
    # 方法1: 画像認識（gui_images/btn_no.png が必要）
    _click_image_or_fallback(
        image_name="btn_no.png",
        fallback_keys=[('n',)],  # フォールバック: N キー（「いいえ」のショートカット）
        timeout=3
    )
    time.sleep(1)
    return True


def set_condition_via_clipboard(condition_string: str) -> bool:
    """クリップボード経由で CPAT に条件文字列を貼り付ける。

    Args:
        condition_string: 貼り付ける文字列

    Returns:
        True: 成功
    """
    import pyperclip
    pyperclip.copy(condition_string)
    if _activate_window("CPAT"):
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.press('enter')
        return True
    return False


# =============================================================================
# gui_images フォルダのセットアップ案内
# =============================================================================

def setup_gui_images():
    """gui_images フォルダを作成し、必要な画像ファイルの一覧を表示する。

    初回セットアップ時にこの関数を実行してください。
    表示されたファイル名でスクリーンショットを撮影し、
    gui_images/ フォルダに保存することで画像認識が有効になります。
    """
    os.makedirs(_IMG_DIR, exist_ok=True)
    print(f"gui_images フォルダを作成しました: {_IMG_DIR}")
    print()
    print("【必要な画像ファイル一覧】")
    print("以下のボタン・ダイアログのスクリーンショットを撮影して")
    print(f"'{_IMG_DIR}' に保存してください。")
    print()

    required_images = [
        ("btn_run.png",      "CPAT の「シミュレーション実行」ボタン"),
        ("btn_save_csv.png", "CPAT の「CSV 保存」ボタン"),
        ("btn_no.png",       "「保存しますか？」ダイアログの「いいえ」ボタン"),
    ]

    for filename, description in required_images:
        path = os.path.join(_IMG_DIR, filename)
        status = "OK（存在します）" if os.path.exists(path) else "未作成"
        print(f"  [{status}] {filename}  ← {description}")

    print()
    print("【スクリーンショットの撮り方】")
    print("  1. CPAT を起動して対象のボタンを画面に表示する")
    print("  2. Snipping Tool（Win+Shift+S）でボタン部分だけを切り抜く")
    print("  3. 上記のファイル名で gui_images/ フォルダに保存する")
    print()
    print("画像ファイルがない場合は、キーボードショートカットによる")
    print("フォールバック操作が自動的に使用されます。")


if __name__ == "__main__":
    # このスクリプトを直接実行すると、セットアップ案内が表示されます
    setup_gui_images()
