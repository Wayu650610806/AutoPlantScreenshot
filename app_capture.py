import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import simpledialog
from PIL import ImageGrab, ImageTk, Image
import threading
import datetime
import sys
import os
import time
import numpy as np
import cv2
import json
import easyocr # For OCR
import shutil # For deleting folders
import requests # For sending data

# --- Base Path Logic ---
def get_base_path():
    """Gets the base path, whether running as .py or frozen .exe"""
    if getattr(sys, 'frozen', False):
        # We are running in a bundle (e.g., PyInstaller)
        base_path = os.path.dirname(sys.executable)
    else:
        # We are running in a normal Python environment
        base_path = os.path.dirname(os.path.abspath(__file__))
    return base_path

# --- Global Paths & Constants ---
# (!!! MUST BE DEFINED *BEFORE* EASYOCR INIT !!!)
BASE_PATH = get_base_path()
MODEL_STORAGE_DIR = os.path.join(BASE_PATH, 'model') # Path for EasyOCR models
TABNAME_DIR = os.path.join(BASE_PATH, "pictures", "tabname") 
STATUS_TEMPLATE_DIR = os.path.join(BASE_PATH, "pictures", "status")
ROI_DIR = os.path.join(BASE_PATH, "rois")
CONFIG_FILE_PATH = os.path.join(BASE_PATH, "config.json")

# Create all necessary folders on startup
os.makedirs(MODEL_STORAGE_DIR, exist_ok=True) 
os.makedirs(TABNAME_DIR, exist_ok=True)
os.makedirs(STATUS_TEMPLATE_DIR, exist_ok=True)
os.makedirs(ROI_DIR, exist_ok=True)


# --- SIFT Global Initialization ---
try:
    sift = cv2.SIFT_create() # (MODIFIED) แก้จาก cv เป็น cv2
    tabname_sift_cache = {}
    status_sift_caches = {}
    # (MODIFIED) นี่คือค่า Default เท่านั้น จะถูกเขียนทับโดย load_config()
    TABNAME_SIFT_THRESHOLD = 70  
    STATUS_SIFT_THRESHOLD = 15   
except Exception as e:
    messagebox.showerror("OpenCV Error", f"ไม่สามารถเริ่ม SIFT ได้ (อาจต้องติดตั้ง opencv-contrib-python)\n{e}")
    sys.exit()

# --- EasyOCR Initialization ---
# (!!! NOW IT CAN FIND MODEL_STORAGE_DIR !!!)
try:
    print("Loading EasyOCR Reader... (This may take a moment on first run)")

    # (!!! MODIFIED !!!) ส่ง path ของ model ที่เรากำหนดเองเข้าไป
    ocr_reader = easyocr.Reader(['en'], model_storage_directory=MODEL_STORAGE_DIR) 

    print("EasyOCR Reader loaded.")
except Exception as e:
    messagebox.showerror("EasyOCR Error", f"Could not initialize EasyOCR.\n{e}")
    sys.exit()
OCR_ALLOWLIST = '-.0123456789'

# --- Predefined ROI Names ---
PREDEFINED_ROI_NAMES = [
    "乾溜ガス化炉A_温度_℃", "乾溜ガス化炉B_温度_℃", "乾溜ガス化炉C_温度_℃",
    "乾溜空気弁A_開度_%", "乾溜空気弁B_開度_%", "乾溜空気弁C_開度_%",
    "乾溜ガス化炉A_運転状況", "乾溜ガス化炉B_運転状況", "乾溜ガス化炉C_運転状況",
    "燃焼炉_温度_℃", "排ガス濃度_CO濃度_ppm", "排ガス濃度_O2濃度_%"
]

# --- (NEW) Predefined lists for Comboboxes ---
PREDEFINED_TABNAME_NAMES = [
    "富山環境整備", "ジェムカ", "ループ", "ニセコ", "武京商会", 
    "光陽建設", "鈴木工業", "直富商事", "九州産廃", "環境整備"
]
PREDEFINED_STATUS_NAMES = [
    "Auto", "None", "Cooling"
]

# --- DPI Awareness Fix ---
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1) 
except ImportError: pass 
except Exception as e: print(f"DPI Awareness Error: {e}")

# --- Split Functions ---
def split_pattern_1(pil_image):
    w, h = pil_image.size
    split_x_1 = (w * 34) // 100; split_x_2 = (w * 68) // 100; split_x_3 = (w * 84) // 100
    return [(0, 0, split_x_1, h), (split_x_1, 0, split_x_2 - split_x_1, h), (split_x_2, 0, split_x_3 - split_x_2, h), (split_x_3, 0, w - split_x_3, h)]
def split_pattern_2(pil_image):
    w, h = pil_image.size
    split_x_1 = w // 4; split_x_2 = w // 2; split_x_3 = (w * 3) // 4
    return [(0, 0, split_x_1, h), (split_x_1, 0, split_x_2 - split_x_1, h), (split_x_2, 0, split_x_3 - split_x_2, h), (split_x_3, 0, w - split_x_3, h)]
def split_pattern_3(pil_image):
    w, h = pil_image.size
    split_x_1 = w // 2; split_x_2 = split_x_1 // 2
    return [(0, 0, split_x_2, h), (split_x_2, 0, split_x_1 - split_x_2, h), (split_x_1, 0, w - split_x_1, h)]
def split_pattern_4(pil_image):
    w, h = pil_image.size
    split_x_1 = w // 2
    return [(0, 0, split_x_1, h), (split_x_1, 0, w - split_x_1, h)]
# --- End Split Functions ---

SPLIT_OPTIONS = {
    "NONE": {'en': 'None (Full Image)', 'ja': 'なし (フルイメージ)', 'func': None},
    "P1_34_34_16_16": {'en': 'Pattern 1 (34/34/16/16)', 'ja': 'パターン1 (34/34/16/16)', 'func': split_pattern_1},
    "P2_25x4": {'en': 'Pattern 2 (25% x 4)', 'ja': 'パターン2 (25% x 4)', 'func': split_pattern_2},
    "P3_25_25_50": {'en': 'Pattern 3 (25/25/50)', 'ja': 'パターン3 (25/25/50)', 'func': split_pattern_3},
    "P4_50_50": {'en': 'Pattern 4 (50 / 50)', 'ja': 'パターン4 (50 / 50)', 'func': split_pattern_4}
}
SPLIT_ORDER = ["NONE", "P1_34_34_16_16", "P2_25x4", "P3_25_25_50", "P4_50_50"]

# --- Language Data (คลังคำศัพท์) ---
translations = {
    'app_title': {'en': 'Capture Tool v0.19 (Minimize)', 'ja': 'キャプチャーツール v0.19 (最小化)'}, # (MODIFIED)
    'tab_capture': {'en': 'Auto-Capture', 'ja': '自動キャプチャ'},
    'tab_gallery': {'en': 'Tabname', 'ja': 'タブ名'}, 
    'tab_roi_sets': {'en': 'ROI Sets', 'ja': 'ROIセット'},
    'tab_status': {'en': 'Status Templates', 'ja': 'ステータス・テンプレート'},
    'tab_settings': {'en': 'Settings', 'ja': '設定'},
    'tab_ocr_debug': {'en': 'OCR Debug', 'ja': 'OCRデバッグ'},
    
    # Capture Tab
    'interval_label': {'en': 'Interval (sec):', 'ja': '間隔 (秒):'},
    'start_button': {'en': 'Start Auto', 'ja': '自動開始'},
    'stop_button': {'en': 'Stop Auto', 'ja': '自動停止'},
    'progress_label': {'en': 'Next capture in:', 'ja': '次のキャプチャ:'},
    'image_placeholder': {'en': 'Captured crops will appear here\n(Press "Start" to begin)', 'ja': 'キャプチャした画像はここに表示されます\n(「開始」を押してください)'},
    'split_method_label': {'en': 'Split Method:', 'ja': '分割方法:'},
    'capture_region_button': {'en': 'Capture Tabname', 'ja': 'タブ名キャプチャ'},
    'minimize_on_start_check': {'en': 'Minimize on Start', 'ja': '開始時に最小化'}, 
    
    # Gallery Tab
    'gallery_preview_header': {'en': 'Preview', 'ja': 'プレビュー'},
    'gallery_refresh_button': {'en': 'Refresh', 'ja': '更新'},
    'gallery_rename_button': {'en': 'Rename', 'ja': '名前変更'},
    'gallery_delete_button': {'en': 'Delete', 'ja': '削除'},
    'gallery_placeholder': {'en': 'Select an image to preview', 'ja': '画像を選択してください'},
    
    # ROI Set Tab
    'create_roi_set_button': {'en': 'Create New ROI Set', 'ja': 'ROIセット新規作成'},
    'add_to_roi_set_button': {'en': 'Add to Selected Set', 'ja': '選択中セットに追加'},
    'roi_set_list_header': {'en': 'ROI Set Files', 'ja': 'ROIセットファイル'},
    'roi_save_as_title': {'en': 'Save ROI Set As', 'ja': 'ROIセットを名前を付けて保存'},
    'roi_save_as_text': {'en': 'Enter a filename for this new ROI set (e.g., "machine_A.json"):', 'ja': 'このROIセットのファイル名を入力してください (例: "machine_A.json"):'},
    'roi_add_another_title': {'en': 'Add ROI', 'ja': 'ROI追加'},
    'roi_add_another_text': {'en': 'ROI "{content}" added to set. Add another one?', 'ja': 'ROI「{content}」をセットに追加しました。続けて追加しますか？'},
    'roi_name_prompt_title': {'en': 'Enter ROI Name', 'ja': 'ROI名入力'},
    'roi_name_prompt_text': {'en': 'Select or type a name for this ROI:', 'ja': 'このROIの名前を選択または入力してください:'},
    'select_roi_set_prompt': {'en': 'Please select an ROI set file first.', 'ja': 'まずROIセットファイルを選択してください。'},
    
    # Status Tab
    'status_folder_header': {'en': 'Tabname Folders', 'ja': 'タブ名フォルダ'},
    'status_image_header': {'en': 'Status Images', 'ja': 'ステータス画像'},
    'create_folder_button': {'en': 'Create Folder', 'ja': 'フォルダ作成'},
    'rename_folder_button': {'en': 'Rename Folder', 'ja': 'フォルダ名変更'},
    'delete_folder_button': {'en': 'Delete Folder', 'ja': 'フォルダ削除'},
    'add_image_button': {'en': 'Add Status Image...', 'ja': 'ステータス画像追加...'},
    'rename_image_button': {'en': 'Rename Image', 'ja': '画像名変更'},
    'delete_image_button': {'en': 'Delete Image', 'ja': '画像削除'},
    'add_image_prompt_title': {'en': 'Enter Status Name', 'ja': 'ステータス名入力'},
    'add_image_prompt_text': {'en': 'Enter name for this status (e.g., "Cooling"):', 'ja': 'このステータスの名前を入力してください (例: "Cooling"):'},
    'select_folder_prompt': {'en': 'Please select a folder first.', 'ja': 'まずフォルダを選択してください。'},
    
    # Settings Tab
    'g_sheet_url_label': {'en': 'Google Sheet Web App URL:', 'ja': 'Google Sheet Web AppのURL:'},
    'g_sheet_save_button': {'en': 'Save All Settings', 'ja': 'すべての設定を保存'},
    'tabname_threshold_label': {'en': 'Tabname SIFT Threshold (e.g., 70):', 'ja': 'タブ名SIFTしきい値 (例: 70):'},
    'status_threshold_label': {'en': 'Status SIFT Threshold (e.g., 15):', 'ja': 'ステータスSIFTしきい値 (例: 15):'},
    
    # (MODIFIED) Settings Tab - OCR Preprocessing
    'ocr_settings_header': {'en': 'OCR Preprocessing Settings (Advanced)', 'ja': 'OCR前処理設定 (詳細)'},
    'ocr_scale_label': {'en': '1. Upscale Factor (e.g., 4):', 'ja': '1. アップスケール係数 (例: 4):'},
    'ocr_scale_help': {'en': 'Larger = slower, but better for tiny text.', 'ja': '大きいほど低速だが、小さい文字に有効。'},
    'ocr_clahe_label': {'en': '2. CLAHE Clip Limit (e.g., 2.0):', 'ja': '2. CLAHEクリップ上限 (例: 2.0):'},
    'ocr_clahe_help': {'en': 'Increases contrast. Lower = less noise (e.g., 1.5). Higher = sharper (e.g., 3.0).', 'ja': 'コントラストを強化。低い = ノイズ減 (例: 1.5)。高い = よりシャープ (例: 3.0)。'},
    'ocr_median_label': {'en': '3. Median Blur ksize (e.g., 3):', 'ja': '3. メディアンブラー ksize (例: 3):'},
    'ocr_median_help': {'en': 'Smooths noise. MUST be an ODD number > 1 (3, 5, 7).', 'ja': 'ノイズを平滑化。1より大きい奇数 (3, 5, 7) である必要があります。'},
    'ocr_opening_label': {'en': '4. Opening Kernel ksize (e.g., 2):', 'ja': '4. オープニングカーネル ksize (例: 2):'},
    'ocr_opening_help': {'en': 'Removes small white dots. Higher = removes more (e.g., 3).', 'ja': '小さい白い点を除去。大きい = より多く除去 (例: 3)。'},
    
    # (NEW) Conditional Morphology Kernel Settings
    'ocr_kernel_settings_header': {'en': 'Morphology Kernel Settings', 'ja': 'モルフォロジー・カーネル設定'},
    'ocr_dilate_label': {'en': '1. Dilate/Thicken Kernel ksize (e.g., 2):', 'ja': '1. 膨張/หนาขึ้น カーネル ksize (例: 2):'}, # Now refers to the Thickening action
    'ocr_dilate_help': {'en': 'Kernel size for thickening dark text (ROI targets set below).', 'ja': '濃い文字を太らせるためのカーネルサイズ (ROIターゲットは下記参照)。'},
    'ocr_erode_label': {'en': '2. Erode/Thin Kernel ksize (e.g., 2):', 'ja': '2. 収縮/บางลง カーネル ksize (例: 2):'}, # Now refers to the Thinning action
    'ocr_erode_help': {'en': 'Kernel size for thinning dark text (ROI targets set below).', 'ja': '濃い文字を細くするためのカーネルサイズ (ROIターゲットは下記参照)。'},
    
    # (NEW) Conditional Morphology Target Selection UI
    'ocr_targets_header': {'en': 'Conditional Morphology Target Selection', 'ja': '条件付き前処理ターゲット選択'},
    'available_roi_label': {'en': 'Available ROI Templates:', 'ja': '利用可能なROIテンプレート:'},
    'dilate_targets_label': {'en': 'Targets for Dilate (Thickening):', 'ja': '膨張 (太らせる) ターゲット:'},
    'erode_targets_label': {'en': 'Targets for Erode (Thinning):', 'ja': '収縮 (細くする) ターゲット:'},
    'add_dilate_button': {'en': '-> Add to Thickening', 'ja': '-> 太らせるに追加'},
    'remove_dilate_button': {'en': '<- Remove', 'ja': '<- 削除'},
    'add_erode_button': {'en': '-> Add to Thinning', 'ja': '-> 細くするに追加'},
    'remove_erode_button': {'en': '<- Remove', 'ja': '<- 削除'},
    'refresh_targets_button': {'en': 'Refresh ROI List', 'ja': 'ROIリストを更新'},
    
    # OCR Debug Tab (MODIFIED for 2x2 grid)
    'ocr_refresh_button': {'en': 'Load Latest Capture Data', 'ja': '最新のキャプチャを読込'},
    'ocr_split_label': {'en': 'Select Captured Split:', 'ja': 'キャプチャ分割を選択:'},
    'ocr_roi_label': {'en': 'Select ROI to Inspect:', 'ja': '検査するROIを選択:'},
    'ocr_raw_label': {'en': 'Raw ROI', 'ja': '生ROI'},
    'ocr_gray_label': {'en': 'Grayscale', 'ja': 'グレースケール'},
    'ocr_contrast_label': {'en': 'Conditional Processing', 'ja': '条件付き前処理'},
    'ocr_final_label': {'en': 'Final (Threshold)', 'ja': '最終 (しきい値処理)'},
    'ocr_result_label': {'en': 'OCR Result:', 'ja': 'OCR結果:'},
    'ocr_no_data': {'en': 'No data. Run Auto-Capture first.', 'ja': 'データなし。自動キャプチャを実行してください。'},
    
    'status_config_saved': {'en': 'Configuration saved.', 'ja': '設定を保存しました。'},
    'rename_prompt_title': {'en': 'Rename File', 'ja': '名前の変更'},
    'rename_prompt_text': {'en': 'Enter new filename:', 'ja': '新しいファイル名を入力:'},
    'delete_confirm_title': {'en': 'Confirm Delete', 'ja': '削除の確認'},
    'delete_confirm_text': {'en': 'Are you sure you want to delete this file?\n{content}', 'ja': 'このファイルを削除してもよろしいですか？\n{content}'},
    'delete_folder_confirm_text': {'en': 'Are you sure you want to PERMANENTLY delete this folder and ALL images inside it?\n{content}', 'ja': 'このフォルダと中の画像をすべて完全に削除してもよろしいですか？\n{content}'},
    'validation_pass': {'en': 'Data OK', 'ja': 'データ正常'},
    'validation_incomplete': {'en': 'Incomplete Data (N/A)', 'ja': 'データ不完全 (N/A)'},
    'validation_invalid': {'en': 'Invalid Data (Out of Range)', 'ja': 'データ異常 (範囲外)'},
    'status_idle': {'en': 'Status: Idle', 'ja': 'ステータス: 待機中'},
    'status_running': {'en': 'Auto-Capture running... (every {content} sec)', 'ja': '自動キャプチャ実行中... ({content} 秒ごと)'},
    'status_stopped': {'en': 'Status: Stopped', 'ja': 'ステータス: 停止'},
    'status_captured': {'en': 'Last captured: {content}', 'ja': '最終キャプチャ: {content}'},
    'status_error': {'en': 'Error: {content}', 'ja': 'エラー: {content}'},
    'status_saved': {'en': 'Image saved to gallery: {content}', 'ja': 'ギャラリーに画像を保存しました: {content}'},
    'status_sift_loading': {'en': 'Loading SIFT templates...', 'ja': 'SIFTテンプレートを読込中...'},
    'status_sift_done': {'en': 'SIFT templates loaded ({content[0]} tabnames, {content[1]} statuses).', 'ja': 'SIFTテンプレートを読込完了 (タブ名{content[0]}件、ステータス{content[1]}件)。'},
    'status_roi_saved': {'en': 'ROI Set "{content}" saved.', 'ja': 'ROIセット「{content}」を保存しました。'},
    'status_roi_error': {'en': 'Failed to load/save ROI data.', 'ja': 'ROIデータの読み込み/保存に失敗しました。'},
    'status_data_sending': {'en': 'Sending data to Google Sheet...', 'ja': 'Google Sheetにデータを送信中...'},
    'status_data_sent': {'en': 'Data sent to sheet: {content}', 'ja': 'シートにデータを送信しました: {content}'},
    'error_title': {'en': 'Invalid Input', 'ja': '無効な入力'},
    'error_message': {'en': 'Please enter a valid number (seconds > 1).\n{content}', 'ja': '有効な数字（1秒以上）を入力してください。\n{content}'},
    'error_threshold': {'en': 'Invalid Threshold', 'ja': '無効なしきい値'},
    'error_threshold_text': {'en': 'Threshold values must be integers.', 'ja': 'しきい値は整数である必要があります。'},
    
    # (NEW) Error messages
    'error_ocr_settings': {'en': 'Invalid OCR Settings', 'ja': '無効なOCR設定'},
    'error_ocr_text': {'en': 'All OCR values must be numbers.\nMedian ksize must be an ODD integer > 1.\nAll other ksize values must be > 0.', 'ja': 'OCR値はすべて数値である必要があります。\nメดิアン ksize は1よりใหญ่ขึ้น奇数である必要があります。\n他のすべてのカーネルサイズ値は0よりใหญ่ขึ้น必要があります。'},
    
    'confirm_close_title': {'en': 'Confirm Exit', 'ja': '終了確認'},
    'confirm_close_message': {'en': 'Auto-Capture is running. Are you sure you want to stop and exit?', 'ja': '自動キャプチャが実行中です。停止して終了しますか？'},
    'lang_button': {'en': '日本語', 'ja': 'English'},
    **{key: value for key, value in SPLIT_OPTIONS.items()} 
}

# --- Global Variables (MODIFIED for 2x2 grid + OCR Settings) ---
is_running = False      
timer_job_id = None     
current_lang = 'en' 
auto_cap_photos = [] 
gallery_preview_photo = None
image_placeholder_label = None
gallery_image_list = None
gallery_preview_label = None
gallery_rename_button = None
gallery_delete_button = None
capture_region_button = None
notebook = None
roi_set_list = None
create_roi_set_button = None
add_to_roi_set_button = None
roi_rename_button = None
roi_delete_button = None
roi_refresh_button = None
status_tab = None
status_folder_list = None
status_image_list = None
status_preview_photo = None
status_preview_label = None
status_add_image_button = None
status_rename_image_button = None
status_delete_image_button = None
status_create_folder_button = None
status_rename_folder_button = None
status_delete_folder_button = None
g_sheet_url = ""
g_sheet_url_entry = None
g_sheet_save_button = None
settings_tab = None
tabname_threshold_entry = None
status_threshold_entry = None
g_latest_sift_results = []
ocr_debug_tab = None
ocr_split_combo = None
ocr_roi_combo = None
ocr_raw_image_label = None
ocr_result_label = None
ocr_raw_photo = None
ocr_gray_frame_label = None 
ocr_gray_image_label = None 
ocr_gray_photo = None 
ocr_contrast_frame_label = None 
ocr_contrast_image_label = None 
ocr_contrast_photo = None 
ocr_final_frame_label = None 
ocr_final_image_label = None 
ocr_final_photo = None 
minimize_on_start_var = None
minimize_on_start_check = None
crop_display_frame = None # (NEW) ทำให้เป็น Global เพื่อให้ clear_image_display รู้จัก

# (NEW) OCR Settings Globals
OCR_SCALE_FACTOR = 4
OCR_CLAHE_CLIP = 2.0
OCR_MEDIAN_KSIZE = 3
OCR_OPENING_KSIZE = 2
OCR_DILATE_KSIZE = 2  
OCR_ERODE_KSIZE = 2   
OCR_DILATE_TARGETS = ["乾溜空気弁A_開度_%", "乾溜空気弁B_開度_%", "乾溜空気弁C_開度_%"] # (MODIFIED Default)
OCR_ERODE_TARGETS = ["燃焼炉_温度_℃"] # (MODIFIED Default)

ocr_scale_entry = None
ocr_clahe_entry = None
ocr_median_entry = None
ocr_opening_entry = None
ocr_dilate_entry = None 
ocr_erode_entry = None  

# (NEW UI Globals for Target Selection)
dilate_target_listbox = None
erode_target_listbox = None
available_roi_listbox = None

# --- Custom Entry with Right-Click ---
class EntryWithRightClickMenu(ttk.Entry):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Cut", command=self.do_cut)
        self.context_menu.add_command(label="Copy", command=self.do_copy)
        self.context_menu.add_command(label="Paste", command=self.do_paste)
        self.bind("<Button-3>", self.show_context_menu)
    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    def do_cut(self): self.event_generate("<<Cut>>")
    def do_copy(self): self.event_generate("<<Copy>>")
    def do_paste(self): self.event_generate("<<Paste>>")

# --- Custom Dialog for ROI Naming ---
class AskROINameDialog(simpledialog.Dialog):
    def __init__(self, parent, title, text, predefined_names, initialvalue=""):
        self.text = text
        self.predefined_names = predefined_names
        self.initialvalue = initialvalue
        super().__init__(parent, title)
    def body(self, master):
        ttk.Label(master, text=self.text).pack(pady=5)
        self.combo = ttk.Combobox(master, values=self.predefined_names, width=50)
        self.combo.pack(padx=10, pady=5)
        self.combo.insert(0, self.initialvalue)
        return self.combo
    def apply(self):
        self.result = self.combo.get()

# --- (NEW) Scrollable Frame Classes (MODIFIED) ---
class ScrollableFrame(ttk.Frame):
    """(MODIFIED) A pure Tkinter scrollable frame with BOTH scrollbars."""
    def __init__(self, parent, *args, **kw):
        ttk.Frame.__init__(self, parent, *args, **kw)

        # 1. Create a canvas
        canvas = tk.Canvas(self, borderwidth=0, background="#f0f0f0", highlightthickness=0)
        
        # 2. Create scrollbars
        vscrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        hscrollbar = ttk.Scrollbar(self, orient="horizontal", command=canvas.xview) # (NEW)
        canvas.configure(yscrollcommand=vscrollbar.set, xscrollcommand=hscrollbar.set) # (MODIFIED)
        
        # 3. Pack them
        vscrollbar.pack(side="right", fill="y")
        hscrollbar.pack(side="bottom", fill="x") # (NEW)
        canvas.pack(side="left", fill="both", expand=True)

        # 4. Create the interior frame
        self.interior = ttk.Frame(canvas, style='TFrame')
        self.interior.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # 5. Pack the interior frame into the canvas
        canvas.create_window((0, 0), window=self.interior, anchor="nw")

class HorizontalScrolledFrame(ttk.Frame):
    """A pure Tkinter horizontal scrollable frame."""
    def __init__(self, parent, *args, **kw):
        ttk.Frame.__init__(self, parent, *args, **kw)

        # 1. Create a canvas
        hscanvas = tk.Canvas(self, borderwidth=0, background="#ffffff", highlightthickness=0)
        
        # 2. Create a scrollbar
        hscrollbar = ttk.Scrollbar(self, orient="horizontal", command=hscanvas.xview)
        hscanvas.configure(xscrollcommand=hscrollbar.set)
        
        # 3. Pack them
        hscrollbar.pack(side="bottom", fill="x")
        hscanvas.pack(side="top", fill="both", expand=True)

        # 4. Create the interior frame
        self.interior = tk.Frame(hscanvas, background="#ffffff")
        self.interior.bind('<Configure>', lambda e: hscanvas.configure(scrollregion=hscanvas.bbox("all")))

        # 5. Pack the interior frame into the canvas
        hscanvas.create_window((0, 0), window=self.interior, anchor="nw")

# --- Config Persistence ---
# --- (MODIFIED) - เพิ่มการโหลด/บันทึก OCR Settings & TARGETS ---
def load_config():
    """(MODIFIED) Loads all settings from config.json."""
    global g_sheet_url, TABNAME_SIFT_THRESHOLD, STATUS_SIFT_THRESHOLD, \
           OCR_SCALE_FACTOR, OCR_CLAHE_CLIP, OCR_MEDIAN_KSIZE, OCR_OPENING_KSIZE, \
           OCR_DILATE_KSIZE, OCR_ERODE_KSIZE, OCR_DILATE_TARGETS, OCR_ERODE_TARGETS # (MODIFIED)
    try:
        if os.path.exists(CONFIG_FILE_PATH):
            with open(CONFIG_FILE_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
                # SIFT/Sheet Settings
                g_sheet_url = data.get("g_sheet_url", "")
                if g_sheet_url_entry:
                    g_sheet_url_entry.delete(0, tk.END)
                    g_sheet_url_entry.insert(0, g_sheet_url)
                
                TABNAME_SIFT_THRESHOLD = int(data.get("tabname_sift_threshold", 70))
                STATUS_SIFT_THRESHOLD = int(data.get("status_sift_threshold", 15))
                if tabname_threshold_entry:
                    tabname_threshold_entry.delete(0, tk.END)
                    tabname_threshold_entry.insert(0, str(TABNAME_SIFT_THRESHOLD))
                if status_threshold_entry:
                    status_threshold_entry.delete(0, tk.END)
                    status_threshold_entry.insert(0, str(STATUS_SIFT_THRESHOLD))
                
                # (NEW) OCR Settings
                OCR_SCALE_FACTOR = int(data.get("ocr_scale_factor", 4))
                OCR_CLAHE_CLIP = float(data.get("ocr_clahe_clip", 2.0))
                OCR_MEDIAN_KSIZE = int(data.get("ocr_median_ksize", 3))
                OCR_OPENING_KSIZE = int(data.get("ocr_opening_ksize", 2))
                
                # (NEW) Conditional Morphology Settings (ksize)
                OCR_DILATE_KSIZE = int(data.get("ocr_dilate_ksize", 2))
                OCR_ERODE_KSIZE = int(data.get("ocr_erode_ksize", 2))
                
                # (NEW) Conditional Morphology Targets
                default_targets_dilate = ["乾溜空気弁A_開度_%", "乾溜空気弁B_開度_%", "乾溜空気弁C_開度_%"]
                default_targets_erode = ["燃焼炉_温度_℃"]
                OCR_DILATE_TARGETS = data.get("ocr_dilate_targets", default_targets_dilate)
                OCR_ERODE_TARGETS = data.get("ocr_erode_targets", default_targets_erode)

                # Update UI Elements
                if ocr_scale_entry:
                    ocr_scale_entry.delete(0, tk.END)
                    ocr_scale_entry.insert(0, str(OCR_SCALE_FACTOR))
                if ocr_clahe_entry:
                    ocr_clahe_entry.delete(0, tk.END)
                    ocr_clahe_entry.insert(0, str(OCR_CLAHE_CLIP))
                if ocr_median_entry:
                    ocr_median_entry.delete(0, tk.END)
                    ocr_median_entry.insert(0, str(OCR_MEDIAN_KSIZE))
                if ocr_opening_entry:
                    ocr_opening_entry.delete(0, tk.END)
                    ocr_opening_entry.insert(0, str(OCR_OPENING_KSIZE))
                if ocr_dilate_entry:
                    ocr_dilate_entry.delete(0, tk.END)
                    ocr_dilate_entry.insert(0, str(OCR_DILATE_KSIZE))
                if ocr_erode_entry:
                    ocr_erode_entry.delete(0, tk.END)
                    ocr_erode_entry.insert(0, str(OCR_ERODE_KSIZE))
                
                # (NEW) Refresh target listboxes if they exist
                if available_roi_listbox:
                    refresh_ocr_target_listboxes()

        else:
            # (NEW) Load defaults into UI if no config file
            if tabname_threshold_entry: tabname_threshold_entry.insert(0, "70")
            if status_threshold_entry: status_threshold_entry.insert(0, "15")
            if g_sheet_url_entry: g_sheet_url_entry.insert(0, "")
            if ocr_scale_entry: ocr_scale_entry.insert(0, "4")
            if ocr_clahe_entry: ocr_clahe_entry.insert(0, "2.0")
            if ocr_median_entry: ocr_median_entry.insert(0, "3")
            if ocr_opening_entry: ocr_opening_entry.insert(0, "2")
            if ocr_dilate_entry: ocr_dilate_entry.insert(0, "2")
            if ocr_erode_entry: ocr_erode_entry.insert(0, "2")
            
    except Exception as e:
        print(f"Error loading config: {e}")
        # (Reset all to default on error)
        g_sheet_url = ""
        TABNAME_SIFT_THRESHOLD = 70
        STATUS_SIFT_THRESHOLD = 15
        OCR_SCALE_FACTOR = 4
        OCR_CLAHE_CLIP = 2.0
        OCR_MEDIAN_KSIZE = 3
        OCR_OPENING_KSIZE = 2
        OCR_DILATE_KSIZE = 2 
        OCR_ERODE_KSIZE = 2  
        OCR_DILATE_TARGETS = ["乾溜空気弁A_開度_%", "乾溜空気弁B_開度_%", "乾溜空気弁C_開度_%"]
        OCR_ERODE_TARGETS = ["燃焼炉_温度_℃"]

def save_config():
    """(MODIFIED) Saves all settings to config.json with validation."""
    global g_sheet_url, TABNAME_SIFT_THRESHOLD, STATUS_SIFT_THRESHOLD, \
           OCR_SCALE_FACTOR, OCR_CLAHE_CLIP, OCR_MEDIAN_KSIZE, OCR_OPENING_KSIZE, \
           OCR_DILATE_KSIZE, OCR_ERODE_KSIZE, OCR_DILATE_TARGETS, OCR_ERODE_TARGETS # (MODIFIED)
    try:
        # 1. Validate SIFT thresholds
        try:
            new_tab_thresh = int(tabname_threshold_entry.get())
            new_stat_thresh = int(status_threshold_entry.get())
        except ValueError:
            messagebox.showerror(translations['error_threshold'][current_lang], translations['error_threshold_text'][current_lang])
            return
            
        # 2. (NEW) Validate OCR Settings
        try:
            new_scale = int(ocr_scale_entry.get())
            new_clahe = float(ocr_clahe_entry.get())
            new_median = int(ocr_median_entry.get())
            new_opening = int(ocr_opening_entry.get())
            
            # (NEW) Conditional Morphology validation (ksize)
            new_dilate = int(ocr_dilate_entry.get())
            new_erode = int(ocr_erode_entry.get())
            
            # Validation rules
            if new_median % 2 == 0 or new_median < 3:
                raise ValueError("Median ksize must be an odd integer > 1")
            if new_scale <= 0 or new_clahe <= 0 or new_opening <= 0:
                 raise ValueError("Values must be > 0")
            # (NEW) Dilate/Erode must be > 0
            if new_dilate <= 0 or new_erode <= 0:
                 raise ValueError("Dilate/Erode ksizes must be > 0")
                 
        except ValueError as e:
            print(f"OCR Setting Validation Error: {e}")
            messagebox.showerror(translations['error_ocr_settings'][current_lang], translations['error_ocr_text'][current_lang])
            return

        # 3. All valid, update Globals
        TABNAME_SIFT_THRESHOLD = new_tab_thresh
        STATUS_SIFT_THRESHOLD = new_stat_thresh
        OCR_SCALE_FACTOR = new_scale
        OCR_CLAHE_CLIP = new_clahe
        OCR_MEDIAN_KSIZE = new_median
        OCR_OPENING_KSIZE = new_opening
        
        # (NEW) Conditional Morphology updates (ksize and targets)
        OCR_DILATE_KSIZE = new_dilate
        OCR_ERODE_KSIZE = new_erode
        
        # Update targets from UI listboxes
        if dilate_target_listbox:
            OCR_DILATE_TARGETS = list(dilate_target_listbox.get(0, tk.END))
        if erode_target_listbox:
            OCR_ERODE_TARGETS = list(erode_target_listbox.get(0, tk.END))

        g_sheet_url = g_sheet_url_entry.get()
        
        # 4. Create data dict
        data = {
            "g_sheet_url": g_sheet_url,
            "tabname_sift_threshold": TABNAME_SIFT_THRESHOLD,
            "status_sift_threshold": STATUS_SIFT_THRESHOLD,
            "ocr_scale_factor": OCR_SCALE_FACTOR,
            "ocr_clahe_clip": OCR_CLAHE_CLIP,
            "ocr_median_ksize": OCR_MEDIAN_KSIZE,
            "ocr_opening_ksize": OCR_OPENING_KSIZE,
            "ocr_dilate_ksize": OCR_DILATE_KSIZE, 
            "ocr_erode_ksize": OCR_ERODE_KSIZE,
            "ocr_dilate_targets": OCR_DILATE_TARGETS, # (NEW)
            "ocr_erode_targets": OCR_ERODE_TARGETS    # (NEW)
        }
        
        # 5. Save to file
        with open(CONFIG_FILE_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        update_status('status_config_saved')
        
    except Exception as e:
        update_status('status_error', f"Config save failed: {e}")

def set_language(lang_code):
    global current_lang, image_placeholder_label
    current_lang = lang_code
    
    root.title(translations['app_title'][current_lang])
    notebook.tab(capture_tab, text=translations['tab_capture'][current_lang])
    notebook.tab(gallery_tab, text=translations['tab_gallery'][current_lang])
    notebook.tab(roi_tab, text=translations['tab_roi_sets'][current_lang])
    notebook.tab(status_tab, text=translations['tab_status'][current_lang])
    notebook.tab(settings_tab_scroller, text=translations['tab_settings'][current_lang]) # (MODIFIED)
    notebook.tab(ocr_debug_tab_scroller, text=translations['tab_ocr_debug'][current_lang]) # (MODIFIED)
    
    # Capture Tab
    interval_label.config(text=translations['interval_label'][current_lang])
    start_button.config(text=translations['start_button'][current_lang])
    stop_button.config(text=translations['stop_button'][current_lang])
    lang_button.config(text=translations['lang_button'][current_lang])
    progress_label.config(text=translations['progress_label'][current_lang])
    split_method_label.config(text=translations['split_method_label'][current_lang])
    capture_region_button.config(text=translations['capture_region_button'][current_lang])
    minimize_on_start_check.config(text=translations['minimize_on_start_check'][current_lang])
    
    current_index = split_method_combo.current()
    if current_index == -1: current_index = 0
    split_display_options = [SPLIT_OPTIONS[key][current_lang] for key in SPLIT_ORDER]
    split_method_combo.config(values=split_display_options)
    split_method_combo.current(current_index)
    
    # Gallery (Tabname) Tab
    gallery_image_list.heading("#0", text="")
    gallery_refresh_button.config(text=translations['gallery_refresh_button'][current_lang])
    gallery_preview_header_label.config(text=translations['gallery_preview_header'][current_lang])
    gallery_rename_button.config(text=translations['gallery_rename_button'][current_lang])
    gallery_delete_button.config(text=translations['gallery_delete_button'][current_lang])
    if not gallery_image_list.selection():
        gallery_preview_label.config(text=translations['gallery_placeholder'][current_lang], image='')
        
    # ROI Set Tab
    roi_set_list.heading("#0", text=translations['roi_set_list_header'][current_lang])
    create_roi_set_button.config(text=translations['create_roi_set_button'][current_lang])
    add_to_roi_set_button.config(text=translations['add_to_roi_set_button'][current_lang])
    roi_rename_button.config(text=translations['gallery_rename_button'][current_lang])
    roi_delete_button.config(text=translations['gallery_delete_button'][current_lang])
    roi_refresh_button.config(text=translations['gallery_refresh_button'][current_lang])
        
    # Status Tab
    status_folder_list.heading("#0", text=translations['status_folder_header'][current_lang])
    status_image_list.heading("#0", text=translations['status_image_header'][current_lang])
    status_create_folder_button.config(text=translations['create_folder_button'][current_lang])
    status_rename_folder_button.config(text=translations['rename_folder_button'][current_lang])
    status_delete_folder_button.config(text=translations['delete_folder_button'][current_lang])
    status_add_image_button.config(text=translations['add_image_button'][current_lang])
    status_rename_image_button.config(text=translations['rename_image_button'][current_lang])
    status_delete_image_button.config(text=translations['delete_image_button'][current_lang])

    # Settings Tab (MODIFIED)
    g_sheet_url_label.config(text=translations['g_sheet_url_label'][current_lang])
    tabname_threshold_label.config(text=translations['tabname_threshold_label'][current_lang])
    status_threshold_label.config(text=translations['status_threshold_label'][current_lang])
    g_sheet_save_button.config(text=translations['g_sheet_save_button'][current_lang])
    
    # (MODIFIED) OCR Settings Labels
    ocr_settings_header.config(text=translations['ocr_settings_header'][current_lang])
    ocr_scale_label.config(text=translations['ocr_scale_label'][current_lang])
    ocr_scale_help.config(text=translations['ocr_scale_help'][current_lang])
    ocr_clahe_label.config(text=translations['ocr_clahe_label'][current_lang])
    ocr_clahe_help.config(text=translations['ocr_clahe_help'][current_lang])
    ocr_median_label.config(text=translations['ocr_median_label'][current_lang])
    ocr_median_help.config(text=translations['ocr_median_help'][current_lang])
    ocr_opening_label.config(text=translations['ocr_opening_label'][current_lang])
    ocr_opening_help.config(text=translations['ocr_opening_help'][current_lang])
    
    ocr_kernel_settings_header.config(text=translations['ocr_kernel_settings_header'][current_lang]) # (NEW)
    ocr_dilate_label.config(text=translations['ocr_dilate_label'][current_lang])
    ocr_dilate_help.config(text=translations['ocr_dilate_help'][current_lang])
    ocr_erode_label.config(text=translations['ocr_erode_label'][current_lang])
    ocr_erode_help.config(text=translations['ocr_erode_help'][current_lang])
    
    ocr_targets_header.config(text=translations['ocr_targets_header'][current_lang]) # (NEW)
    available_roi_label.config(text=translations['available_roi_label'][current_lang]) # (NEW)
    dilate_targets_label.config(text=translations['dilate_targets_label'][current_lang]) # (NEW)
    erode_targets_label.config(text=translations['erode_targets_label'][current_lang]) # (NEW)
    add_dilate_button.config(text=translations['add_dilate_button'][current_lang]) # (NEW)
    remove_dilate_button.config(text=translations['remove_dilate_button'][current_lang]) # (NEW)
    add_erode_button.config(text=translations['add_erode_button'][current_lang]) # (NEW)
    remove_erode_button.config(text=translations['remove_erode_button'][current_lang]) # (NEW)
    refresh_targets_button.config(text=translations['refresh_targets_button'][current_lang]) # (NEW)
    
    # OCR Debug Tab (MODIFIED for 2x2 grid)
    ocr_refresh_btn.config(text=translations['ocr_refresh_button'][current_lang])
    ocr_split_label.config(text=translations['ocr_split_label'][current_lang])
    ocr_roi_label.config(text=translations['ocr_roi_label'][current_lang])
    ocr_raw_frame_label.config(text=translations['ocr_raw_label'][current_lang])
    ocr_gray_frame_label.config(text=translations['ocr_gray_label'][current_lang]) 
    ocr_contrast_frame_label.config(text=translations['ocr_contrast_label'][current_lang]) 
    ocr_final_frame_label.config(text=translations['ocr_final_label'][current_lang]) 
    ocr_result_text_label.config(text=translations['ocr_result_label'][current_lang])
    
    if not is_running:
        status_label.config(text=translations['status_idle'][current_lang])
        if image_placeholder_label:
            image_placeholder_label.config(text=translations['image_placeholder'][current_lang])
    else:
        interval = interval_entry.get()
        update_status('status_running', interval)
    
    refresh_ocr_target_listboxes() # (NEW) Ensure listboxes are refreshed with correct language/data

def toggle_language():
    if current_lang == 'en': set_language('ja')
    else: set_language('en')

# --- Helper functions for Target List Management ---
def refresh_ocr_target_listboxes():
    """Populates the available ROI listbox and selected lists."""
    global OCR_DILATE_TARGETS, OCR_ERODE_TARGETS

    if not available_roi_listbox or not dilate_target_listbox or not erode_target_listbox:
        return

    # 1. Clear all lists
    available_roi_listbox.delete(0, tk.END)
    dilate_target_listbox.delete(0, tk.END)
    erode_target_listbox.delete(0, tk.END)

    # 2. Ensure targets are unique (cannot be in both Dilate and Erode)
    all_targets = set(OCR_DILATE_TARGETS) | set(OCR_ERODE_TARGETS)
    
    # 3. Separate targets from non-targets
    available_rois = [name for name in PREDEFINED_ROI_NAMES if name not in all_targets]
    
    # Re-filter the global targets based on PREDEFINED list to keep the UI clean
    OCR_DILATE_TARGETS = sorted([name for name in OCR_DILATE_TARGETS if name in PREDEFINED_ROI_NAMES])
    OCR_ERODE_TARGETS = sorted([name for name in OCR_ERODE_TARGETS if name in PREDEFINED_ROI_NAMES])

    # 4. Populate lists
    for name in sorted(available_rois):
        available_roi_listbox.insert(tk.END, name)
    for name in OCR_DILATE_TARGETS:
        dilate_target_listbox.insert(tk.END, name)
    for name in OCR_ERODE_TARGETS:
        erode_target_listbox.insert(tk.END, name)

def _move_roi_item(source_listbox, target_type):
    """Handles moving selected items between listboxes."""
    global OCR_DILATE_TARGETS, OCR_ERODE_TARGETS
    
    selected_indices = source_listbox.curselection()
    if not selected_indices: return

    selected_items = [source_listbox.get(i) for i in selected_indices]

    if source_listbox == available_roi_listbox:
        # Moving FROM available TO target list
        if target_type == 'dilate':
            # Remove from all other lists
            OCR_ERODE_TARGETS = [name for name in OCR_ERODE_TARGETS if name not in selected_items]
            OCR_DILATE_TARGETS.extend(selected_items)
        elif target_type == 'erode':
            # Remove from all other lists
            OCR_DILATE_TARGETS = [name for name in OCR_DILATE_TARGETS if name not in selected_items]
            OCR_ERODE_TARGETS.extend(selected_items)
    else:
        # Moving FROM target list TO available
        if target_type == 'dilate':
            OCR_DILATE_TARGETS = [name for name in OCR_DILATE_TARGETS if name not in selected_items]
        elif target_type == 'erode':
            OCR_ERODE_TARGETS = [name for name in OCR_ERODE_TARGETS if name not in selected_items]

    # Re-sort and refresh lists (simple solution)
    refresh_ocr_target_listboxes()

def add_dilate_target():
    _move_roi_item(available_roi_listbox, 'dilate')
def remove_dilate_target():
    _move_roi_item(dilate_target_listbox, 'dilate')
def add_erode_target():
    _move_roi_item(available_roi_listbox, 'erode')
def remove_erode_target():
    _move_roi_item(erode_target_listbox, 'erode')

# --- SIFT & OCR Processing Logic ---
def pil_to_cv2_gray(pil_image):
    return cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2GRAY)

def _load_sift_from_file(filepath):
    try:
        img_bytes = np.fromfile(filepath, dtype=np.uint8)
        img = cv2.imdecode(img_bytes, cv2.IMREAD_GRAYSCALE)
        if img is None: return (None, None)
        kp, des = sift.detectAndCompute(img, None)
        if des is not None and len(kp) > 0:
            return (kp, des)
    except Exception as e:
        print(f"Error loading SIFT from {filepath}: {e}")
    return (None, None)

def load_all_sift_templates():
    """Loads BOTH Tabname and Status templates."""
    global tabname_sift_cache, status_sift_caches
    tabname_sift_cache.clear()
    status_sift_caches.clear()
    update_status('status_sift_loading')
    root.update_idletasks()
    tabname_count = 0
    status_count = 0
    try:
        for filename in os.listdir(TABNAME_DIR):
            if filename.endswith('.png'):
                filepath = os.path.join(TABNAME_DIR, filename)
                kp, des = _load_sift_from_file(filepath)
                if kp:
                    tabname_sift_cache[filename] = (kp, des)
                    tabname_count += 1
    except Exception as e:
        update_status('status_error', f"Tabname SIFT load failed: {e}")
    try:
        for tabname_folder in os.listdir(STATUS_TEMPLATE_DIR):
            tabname_key = tabname_folder + ".png"
            sub_folder_path = os.path.join(STATUS_TEMPLATE_DIR, tabname_folder)
            if os.path.isdir(sub_folder_path):
                status_sift_caches[tabname_key] = {}
                for status_filename in os.listdir(sub_folder_path):
                    if status_filename.endswith('.png'):
                        filepath = os.path.join(sub_folder_path, status_filename)
                        kp, des = _load_sift_from_file(filepath)
                        if kp:
                            status_sift_caches[tabname_key][status_filename] = (kp, des)
                            status_count += 1
        update_status('status_sift_done', (tabname_count, status_count))
    except Exception as e:
        update_status('status_error', f"Status SIFT load failed: {e}")

def _find_best_sift_match(image_to_check_pil, template_cache, match_threshold):
    if not template_cache: return "None"
    try:
        img_crop_gray = pil_to_cv2_gray(image_to_check_pil)
        kp_crop, des_crop = sift.detectAndCompute(img_crop_gray, None)
        if des_crop is None or len(kp_crop) < match_threshold:
            return "None"
        FLANN_INDEX_KDTREE = 1
        index_params = dict(algorithm=FLANN_INDEX_KDTREE, trees=5)
        search_params = dict(checks=50)
        flann = cv2.FlannBasedMatcher(index_params, search_params)
        best_match_name = "None"
        max_good_matches = 0
        for filename, (kp_template, des_template) in template_cache.items():
            if des_template is None: continue
            if len(kp_crop) < 2 or len(kp_template) < 2: continue
            matches = flann.knnMatch(des_crop, des_template, k=2)
            good_matches = [m for m, n in matches if m.distance < 0.7 * n.distance]
            if len(good_matches) > match_threshold and len(good_matches) > max_good_matches:
                max_good_matches = len(good_matches)
                best_match_name = filename
        return best_match_name
    except Exception as e:
        print(f"SIFT match error: {e}")
        return "None"

def find_best_tabname_match(cropped_pil_image):
    w, h = cropped_pil_image.size
    top_half_box = (0, 0, w, h // 2)
    image_to_check = cropped_pil_image.crop(top_half_box)
    return _find_best_sift_match(image_to_check, tabname_sift_cache, TABNAME_SIFT_THRESHOLD)

def find_best_status_match(roi_crop_pil, tabname_match_key):
    if tabname_match_key not in status_sift_caches:
        return "None"
    status_cache_for_this_tab = status_sift_caches[tabname_match_key]
    return _find_best_sift_match(roi_crop_pil, status_cache_for_this_tab, STATUS_SIFT_THRESHOLD)

# --- (MODIFIED) - ใช้ Global Variables จาก Config ---
def preprocess_for_ocr(pil_image, roi_key, scale_factor=None): # (MODIFIED) added roi_key
    """
    (MODIFIED) Uses GLOBAL variables from config for processing.
    (FIXED) Uses target lists for conditional morphology.
    """
    global OCR_DILATE_TARGETS, OCR_ERODE_TARGETS, OCR_DILATE_KSIZE, OCR_ERODE_KSIZE # Use global targets

    try:
        cv_img = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
        
        # 1. Upscale (Uses global OCR_SCALE_FACTOR)
        width = int(cv_img.shape[1] * OCR_SCALE_FACTOR)
        height = int(cv_img.shape[0] * OCR_SCALE_FACTOR)
        if width == 0 or height == 0: return None
        upscaled = cv2.resize(cv_img, (width, height), interpolation=cv2.INTER_LANCZOS4)
        
        # 2. Grayscale
        gray = cv2.cvtColor(upscaled, cv2.COLOR_BGR2GRAY)
        
        # 3. Contrast Enhancement (CLAHE) (Uses global OCR_CLAHE_CLIP)
        clahe = cv2.createCLAHE(clipLimit=OCR_CLAHE_CLIP, tileGridSize=(8,8))
        contrast = clahe.apply(gray) 
        
        # 4. Denoise (Uses global OCR_MEDIAN_KSIZE)
        denoised = cv2.medianBlur(contrast, OCR_MEDIAN_KSIZE) 
        
        # --- (FIXED) CONDITIONAL MORPHOLOGY (Use target lists and swapped logic for dark foreground) ---
        conditional_img = denoised.copy()

        # 4.1. Thickening (Intended Dilate for "開度"): Use ERODE on grayscale if ROI is a DILATE target
        if roi_key in OCR_DILATE_TARGETS:
            # Use kernel size from OCR_DILATE_KSIZE (intended for 'thickening')
            kernel_erode = np.ones((OCR_DILATE_KSIZE, OCR_DILATE_KSIZE), np.uint8) 
            # Use ERODE operation (thickens dark text)
            conditional_img = cv2.erode(conditional_img, kernel_erode, iterations=1) 

        # 4.2. Thinning (Intended Erode for "燃焼炉_温度_℃"): Use DILATE on grayscale if ROI is an ERODE target
        elif roi_key in OCR_ERODE_TARGETS:
            # Use kernel size from OCR_ERODE_KSIZE (intended for 'thinning')
            kernel_dilate = np.ones((OCR_ERODE_KSIZE, OCR_ERODE_KSIZE), np.uint8) 
            # Use DILATE operation (thins dark text)
            conditional_img = cv2.dilate(conditional_img, kernel_dilate, iterations=1)
        
        # --- END CONDITIONAL MORPHOLOGY ---
        
        # 5. Global Thresholding (Otsu's Binarization) - Applied to the conditional_img
        _ , bw_img = cv2.threshold(conditional_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # 6. Invert
        bw_img_inverted = cv2.bitwise_not(bw_img)
        
        # 7. Post-Process Cleaning (Uses global OCR_OPENING_KSIZE)
        kernel_opening = np.ones((OCR_OPENING_KSIZE, OCR_OPENING_KSIZE), np.uint8) 
        final_cleaned = cv2.morphologyEx(bw_img_inverted, cv2.MORPH_OPEN, kernel_opening, iterations=1)

        # Return dict for debug tab
        return {
            'raw_pil': pil_image,
            'gray': gray,
            'contrast': conditional_img, # Now holds the image after conditional morphology
            'final': final_cleaned 
        }
    except Exception as e:
        print(f"OCR Preprocessing error: {e}")
        return None

# --- Data Validation Logic ---
def validate_data(data_results):
    """
    Checks if data is complete (no 'N/A') and valid (follows rules).
    Returns a status string and color.
    """
    if not data_results:
        return "", "black"
    is_complete = True
    is_valid = True
    
    # --- (START) นี่คือจุดที่แก้ไข ---
    # เพิ่ม "Corrupt ROI" เข้าไปใน list ของค่าที่ถือว่าไม่สมบูรณ์
    for key, value in data_results.items():
        if value in ["N/A", "Error", "Corrupt ROI"]: 
            is_complete = False
            break
    # --- (END) นี่คือจุดที่แก้ไข ---
            
    if not is_complete:
        return translations['validation_incomplete'][current_lang], "red"
        
    for key, value in data_results.items():
        try:
            num_val = float(value)
            if ("℃" in key or "ppm" in key) and num_val < 0:
                is_valid = False
                break
            elif ("%" in key) and not (0 <= num_val <= 100):
                is_valid = False
                break
        except (ValueError, TypeError):
            pass 
            
    if not is_valid:
        return translations['validation_invalid'][current_lang], "orange"
        
    return translations['validation_pass'][current_lang], "green"

# --- Google Sheet Upload Logic ---
def _send_data_worker(url, payload):
    """Worker thread to send data without freezing GUI."""
    try:
        root.after(0, update_status, 'status_data_sending')
        response = requests.post(url, json=payload, timeout=10)
        response.raise_for_status()
        root.after(0, update_status, 'status_data_sent', payload.get("sheetName"))
    except requests.RequestException as e:
        root.after(0, update_status, 'status_error', f"GSheet: {e}")

def send_data_to_google_sheet(tabname, data_results):
    """Formats data and starts the sender thread."""
    if not g_sheet_url:
        return
    try:
        sheetName = tabname.replace(".png", "")
        headers = ["Timestamp"] + list(data_results.keys())
        values = [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")] + list(data_results.values())
        payload = {
            "sheetName": sheetName,
            "headers": headers,
            "values": values
        }
        threading.Thread(target=_send_data_worker, args=(g_sheet_url, payload), daemon=True).start()
    except Exception as e:
        root.after(0, update_status, 'status_error', f"GSheet formatting: {e}")

# --- Auto-Capture Logic ---
def start_capture():
    global is_running, timer_job_id
    if is_running: return
    try:
        interval = int(interval_entry.get())
        if interval <= 1: raise ValueError("Time must be > 1")
    except ValueError as e:
        messagebox.showerror(translations['error_title'][current_lang], translations['error_message'][current_lang].format(content=e))
        return
    if minimize_on_start_var.get():
        root.iconify()
        root.update()
        time.sleep(0.5)
    is_running = True
    update_status('status_running', interval)
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)
    interval_entry.config(state=tk.DISABLED)
    lang_button.config(state=tk.DISABLED)
    split_method_combo.config(state=tk.DISABLED)
    capture_region_button.config(state=tk.DISABLED)
    perform_capture_task() 
    update_countdown(0, interval)

def stop_capture():
    global is_running, timer_job_id
    if timer_job_id:
        root.after_cancel(timer_job_id) 
        timer_job_id = None
    is_running = False
    start_button.config(state=tk.NORMAL)
    stop_button.config(state=tk.DISABLED)
    interval_entry.config(state=tk.NORMAL)
    lang_button.config(state=tk.NORMAL)
    split_method_combo.config(state=tk.NORMAL)
    capture_region_button.config(state=tk.NORMAL)
    progress_bar['value'] = 0
    update_status('status_stopped')
    clear_image_display() 

def update_countdown(current_step, interval):
    global timer_job_id
    if not is_running: return
    progress = (current_step / interval) * 100
    progress_bar['value'] = progress
    if current_step < interval:
        timer_job_id = root.after(1000, update_countdown, current_step + 1, interval)
    else:
        progress_bar['value'] = 0 
        threading.Thread(target=perform_capture_task, daemon=True).start()
        timer_job_id = root.after(100, update_countdown, 0, interval) 

def perform_capture_task():
    """(MODIFIED v0.16) Calls validation and GSheet sender."""
    global g_latest_sift_results # (NEW)
    try:
        # (NEW) If minimizing, wait an extra moment
        if minimize_on_start_var.get():
            time.sleep(0.5) # Give 0.5s for window to minimize before capture
            
        image = ImageGrab.grab(all_screens=True)
        selected_index = split_method_combo.current()
        method_key = SPLIT_ORDER[selected_index]
        split_function = SPLIT_OPTIONS[method_key]['func']
        
        final_results = []
        
        if split_function:
            boxes = split_function(image)
            for (x, y, w, h) in boxes:
                box_pil_coords = (x, y, x + w, y + h)
                crop_pil = image.crop(box_pil_coords)
                match_name = find_best_tabname_match(crop_pil)
                data_results = {}
                if match_name != "None":
                    data_results = extract_data_from_rois(crop_pil, match_name, (x, y))
                
                # (NEW) Validate data
                status_text, status_color = validate_data(data_results)
                
                # (NEW) Send data if valid
                if match_name != "None" and status_color == "green":
                    send_data_to_google_sheet(match_name, data_results)
                    
                final_results.append((crop_pil, match_name, (x, y), data_results, (status_text, status_color)))
        else:
            # --- (THIS IS THE FIX FOR FULL SCREEN) ---
            crop_pil = image
            match_name = find_best_tabname_match(crop_pil)
            data_results = {}
            if match_name != "None":
                data_results = extract_data_from_rois(crop_pil, match_name, (0, 0))
            
            # (NEW) Validate data
            status_text, status_color = validate_data(data_results)
            
            # (NEW) Send data if valid
            if match_name != "None" and status_color == "green":
                send_data_to_google_sheet(match_name, data_results)
                
            # (THE FIX) Ensure the 5-tuple is created here too
            final_results.append((crop_pil, match_name, (0, 0), data_results, (status_text, status_color)))

        g_latest_sift_results = final_results # (NEW) Save for debug tab
        root.after(0, update_gui_with_sift_results, final_results)
    except Exception as e:
        root.after(0, update_status, 'status_error', str(e))

def extract_data_from_rois(pil_image, tabname_match, crop_offset):
    """SIFT for '運転状況', Upscaled OCR for ALL OTHERS."""
    data_results = {}
    crop_offset_x, crop_offset_y = crop_offset
    roi_filename = tabname_match.replace(".png", "") + ".json"
    roi_filepath = os.path.join(ROI_DIR, roi_filename)
    
    if not os.path.exists(roi_filepath):
        return {}
    try:
        with open(roi_filepath, 'r', encoding='utf-8') as f:
            rois_to_draw = json.load(f)
    except Exception as e:
        print(f"Error loading ROI file {roi_filename}: {e}")
        return {}
        
    # --- (START) นี่คือจุดที่แก้ไข ---
    for roi_key, roi_value in rois_to_draw.items():
        try:
            # 1. ตรวจสอบว่าข้อมูลใน .json ถูกต้องหรือไม่ (ต้องเป็น list หรือ tuple ที่มี 4 ค่า)
            if not isinstance(roi_value, (list, tuple)) or len(roi_value) != 4:
                print(f"Skipping corrupted ROI '{roi_key}' in {roi_filename}: Expected 4 coordinates, got {roi_value}")
                data_results[roi_key] = "Corrupt ROI"
                continue # ข้าม ROI ที่เสียนี้ไป

            # 2. ถ้าถูกต้อง ก็ unpack ตามปกติ
            global_x, global_y, global_w, global_h = roi_value
            # --- (END) นี่คือจุดที่แก้ไข ---

            local_x = global_x - crop_offset_x
            local_y = global_y - crop_offset_y
            roi_box_pil = (local_x, local_y, local_x + global_w, local_y + global_h)
            
            img_w, img_h = pil_image.size
            if local_x > img_w or local_y > img_h: continue

            roi_crop_pil = pil_image.crop(roi_box_pil)

            if "運転状況" in roi_key:
                status_match = find_best_status_match(roi_crop_pil, tabname_match)
                data_results[roi_key] = status_match.replace(".png", "")
            else:
                # (MODIFIED) Pass roi_key to preprocess_for_ocr
                processing_steps = preprocess_for_ocr(roi_crop_pil, roi_key) 
                if processing_steps is None:
                    data_results[roi_key] = "N/A"
                    continue
                
                # ใช้ภาพ 'final' ในการส่งให้ OCR
                final_processed_img = processing_steps['final'] 
                
                ocr_results = ocr_reader.readtext(final_processed_img, allowlist=OCR_ALLOWLIST, detail=0)
                extracted_text = "".join(ocr_results).strip()
                if not extracted_text:
                    data_results[roi_key] = "N/A"
                else:
                    data_results[roi_key] = extracted_text
            
        except Exception as e:
            print(f"Error processing ROI {roi_key}: {e}")
            data_results[roi_key] = "Error"
    return data_results

def clear_image_display():
    global image_placeholder_label, auto_cap_photos, crop_display_frame
    auto_cap_photos.clear()
    
    # (MODIFIED) - ต้องอ้างอิง crop_display_frame (frame ด้านใน scroller)
    if crop_display_frame:
        for widget in crop_display_frame.winfo_children():
            widget.destroy()
    
    image_placeholder_label = tk.Label(crop_display_frame, font=(font_family, 12, 'italic'), background='#ffffff', foreground='#888888', text=translations['image_placeholder'][current_lang])
    image_placeholder_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)


def update_gui_with_sift_results(sift_results):
    """(REPLACED v0.13) New side-by-side layout."""
    global auto_cap_photos, image_placeholder_label, crop_display_frame
    
    auto_cap_photos.clear()
    
    # (MODIFIED) - ต้องอ้างอิง crop_display_frame (frame ด้านใน scroller)
    for widget in crop_display_frame.winfo_children():
        widget.destroy()
        
    image_placeholder_label = None 
    if not sift_results:
        clear_image_display()
        return

    num_images = len(sift_results)
    
    # (THE FIX) This loop now correctly expects 5 items
    for (pil_image, match_name, (crop_offset_x, crop_offset_y), data_results, (status_text, status_color)) in sift_results:
        
        # (MODIFIED) - pack ลงใน crop_display_frame
        result_frame = tk.Frame(crop_display_frame, background="#f0f0f0", relief=tk.SUNKEN, borderwidth=1)
        
        # 2. Left frame for Image
        image_frame = ttk.Frame(result_frame)
        image_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 10), pady=5)
        
        # 3. Right frame for Data
        data_frame = ttk.Frame(result_frame, width=250)
        data_frame.pack(side=tk.LEFT, fill=tk.Y, pady=5, ipadx=10)
        # ลบ data_frame.pack_propagate(False) ทิ้งไปเลย
        # เพื่อให้ Frame ขยายความสูงตามข้อความได้อิสระ

        # --- Populate Image Frame ---
        display_image = pil_image.copy()
        if match_name != "None":
            try:
                cv_image = cv2.cvtColor(np.array(display_image), cv2.COLOR_RGB2BGR)
                roi_filename = match_name.replace(".png", "") + ".json"
                roi_filepath = os.path.join(ROI_DIR, roi_filename)
                
                if os.path.exists(roi_filepath):
                    with open(roi_filepath, 'r', encoding='utf-8') as f:
                        rois_to_draw = json.load(f)
                    
                    for roi_key, [global_x, global_y, global_w, global_h] in rois_to_draw.items():
                        local_x = global_x - crop_offset_x
                        local_y = global_y - crop_offset_y
                        img_h, img_w = cv_image.shape[:2]
                        if local_x > img_w or local_y > img_h or (local_x + global_w) < 0 or (local_y + global_h) < 0:
                            continue
                        
                        color = (0, 0, 255) # Red (OCR)
                        if "運転状況" in roi_key:
                            color = (255, 0, 0) # Blue (SIFT)
                        cv2.rectangle(cv_image, (local_x, local_y), (local_x + global_w, local_y + global_h), color, 2)
                
                display_image = Image.fromarray(cv2.cvtColor(cv_image, cv2.COLOR_BGR2RGB))
            except Exception as e:
                print(f"Error drawing ROI: {e}")

        # (MODIFIED) - คำนวณความกว้างแบบคงที่ขึ้นต่ำ เพื่อให้ scroller ทำงาน
        min_img_width = 400
        container_width = max(min_img_width, (root.winfo_width() // num_images) - 300)
        display_image.thumbnail((container_width, 500), Image.Resampling.LANCZOS)
        
        photo = ImageTk.PhotoImage(display_image)
        auto_cap_photos.append(photo)
        
        image_label = tk.Label(image_frame, image=photo, background="#ffffff")
        image_label.image = photo 
        image_label.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- Populate Data Frame ---
        display_name = match_name.replace(".png", "")
        name_color = "green" if match_name != "None" else "red"
        
        name_label = ttk.Label(data_frame, text=display_name, font=(font_family, 11, 'bold'), foreground=name_color, anchor=tk.W)
        name_label.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        
        validation_label = ttk.Label(data_frame, text=status_text, font=(font_family, 10, 'bold'), foreground=status_color, anchor=tk.W)
        validation_label.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        data_text = "\n".join([f"{key}: {value}" for key, value in data_results.items()])
        if not data_text:
            data_text = "No ROI data found."
            
        data_label = ttk.Label(data_frame, text=data_text, font=(font_family, 9), 
                              foreground="black", justify=tk.LEFT, anchor=tk.NW)
        data_label.pack(side=tk.TOP, fill=tk.X)

        result_frame.pack(side=tk.LEFT, fill=tk.NONE, expand=False, padx=2, pady=2) # (MODIFIED) ไม่ expand

    now_time = datetime.datetime.now().strftime('%H:%M:%S')
    update_status('status_captured', now_time)


# --- Region Capture & ROI Logic ---
# (!!! START: นี่คือคลาสที่แก้ไขทั้งหมด !!!)
class RegionSelector:
    def __init__(self, parent):
        self.parent = parent
        self.background_image = ImageGrab.grab(all_screens=True)
        self.selector_window = tk.Toplevel(parent)
        
        # --- (START) นี่คือจุดที่แก้ไข ---
        # เราจะบังคับให้หน้าต่างขยายเต็มทุกจอเอง โดยไม่ใช้ -fullscreen
        try:
            # ใช้ windll ที่ import ไว้แล้วที่ด้านบนของไฟล์
            SM_XVIRTUALSCREEN = 76
            SM_YVIRTUALSCREEN = 77
            SM_CXVIRTUALSCREEN = 78
            SM_CYVIRTUALSCREEN = 79

            screen_x = windll.user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
            screen_y = windll.user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
            screen_width = windll.user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
            screen_height = windll.user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
            
            # สร้าง geometry string: "widthxheight+x+y"
            # (x, y อาจเป็นค่าติดลบ ถ้าจอรองอยู่ด้านซ้ายของจอหลัก)
            geometry_string = f"{screen_width}x{screen_height}+{screen_x}+{screen_y}"
            
            self.selector_window.geometry(geometry_string)
            self.selector_window.overrideredirect(True) # ลบขอบหน้าต่างและ title bar
        except Exception as e:
            print(f"Virtual screen detection failed, falling back to primary: {e}")
            self.selector_window.attributes('-fullscreen', True) # ใช้วิธีเดิมถ้าพลาด
        # --- (END) นี่คือจุดที่แก้ไข ---

        self.selector_window.attributes('-alpha', 0.3)
        
        # (แก้ไข) เพิ่ม borderwidth=0 และ highlightthickness=0
        self.canvas = tk.Canvas(self.selector_window, cursor="cross", borderwidth=0, highlightthickness=0) 
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.tk_background_image = ImageTk.PhotoImage(self.background_image)
        
        # ภาพพื้นหลังจะถูกวาดที่ (0, 0) ของ canvas
        # ซึ่งตอนนี้ canvas ก็มีขนาดเท่ากับ virtual desktop ทั้งหมดแล้ว
        self.canvas.create_image(0, 0, image=self.tk_background_image, anchor=tk.NW)
        
        self.rect = None
        self.start_x = 0
        self.start_y = 0
        self.box = None
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_motion)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

    def on_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if not self.rect:
            self.rect = self.canvas.create_rectangle(0, 0, 1, 1, outline='red', width=2, fill='white')
            self.canvas.itemconfigure(self.rect, stipple="gray50") 
            
    def on_motion(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)
        
    def on_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        
        # พิกัด box นี้จะสัมพันธ์กับ (0,0) ของ canvas 
        # ซึ่งก็คือ (0,0) ของ background_image พอดี
        self.box = (int(min(self.start_x, end_x)), int(min(self.start_y, end_y)), int(max(self.start_x, end_x)), int(max(self.start_y, end_y)))
        self.selector_window.destroy()
# (!!! END: สิ้นสุดคลาสที่แก้ไข !!!)


def start_region_capture():
    """Called by 'Capture Tabname' button."""
    root.withdraw()
    time.sleep(0.5)
    selector = RegionSelector(root)
    root.wait_window(selector.selector_window)
    root.deiconify()
    if selector.box:
        try:
            cropped_image = selector.background_image.crop(selector.box)
            dialog = AskROINameDialog(root,
                translations['capture_region_button'][current_lang],
                translations['rename_prompt_text'][current_lang],
                PREDEFINED_TABNAME_NAMES
            )
            filename = dialog.result
            if filename: filename = filename.strip()
            if not filename: return
            if not filename.endswith('.png'):
                filename += '.png'
            save_path = os.path.join(TABNAME_DIR, filename)
            if os.path.exists(save_path):
                if not messagebox.askyesno("Confirm Overwrite", f"File '{filename}' already exists. Overwrite?"):
                    return
            cropped_image.save(save_path)
            update_status('status_saved', filename)
            refresh_gallery_list()
        except Exception as e:
            update_status('status_error', str(e))

def _roi_creation_loop(roi_data_dict):
    """Helper loop for creating/adding ROIs."""
    while True:
        root.withdraw()
        time.sleep(0.5)
        selector = RegionSelector(root)
        root.wait_window(selector.selector_window)
        root.deiconify()
        if not selector.box:
            break
        dialog = AskROINameDialog(root, 
            translations['roi_name_prompt_title'][current_lang],
            translations['roi_name_prompt_text'][current_lang],
            PREDEFINED_ROI_NAMES
        )
        roi_name = dialog.result
        if roi_name:
            x = selector.box[0]
            y = selector.box[1]
            w = selector.box[2] - x
            h = selector.box[3] - y
            roi_data_dict[roi_name] = [x, y, w, h]
            if not messagebox.askyesno(
                translations['roi_add_another_title'][current_lang],
                translations['roi_add_another_text'][current_lang].format(content=roi_name)
            ):
                break
        else:
            if not messagebox.askyesno("Cancel?", "No name entered. Stop this process?"):
                continue
            else:
                break
    return roi_data_dict

def start_roi_set_creation():
    """Called by 'Create New ROI Set' button."""
    new_rois = _roi_creation_loop({})
    if not new_rois:
        return
    dialog = AskROINameDialog(root,
        translations['roi_save_as_title'][current_lang],
        translations['roi_save_as_text'][current_lang],
        PREDEFINED_TABNAME_NAMES
    )
    set_filename = dialog.result
    if set_filename: set_filename = set_filename.strip()
    if set_filename:
        if not set_filename.endswith('.json'):
            set_filename += '.json'
        save_path = os.path.join(ROI_DIR, set_filename)
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(new_rois, f, indent=4, ensure_ascii=False)
            update_status('status_roi_saved', set_filename)
            refresh_roi_file_list()
        except Exception as e:
            update_status('status_error', str(e))

def start_add_to_roi_set():
    """Called by 'Add to Selected Set' button."""
    selected_items = roi_set_list.selection()
    if not selected_items:
        messagebox.showwarning("No Selection", translations['select_roi_set_prompt'][current_lang])
        return
    set_filename = selected_items[0]
    filepath = os.path.join(ROI_DIR, set_filename)
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            existing_rois = json.load(f)
    except Exception as e:
        existing_rois = {} # (FIX) If file is empty, start fresh
    
    updated_rois = _roi_creation_loop(existing_rois)
    
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(updated_rois, f, indent=4, ensure_ascii=False)
        update_status('status_roi_saved', set_filename)
    except Exception as e:
        update_status('status_error', str(e))

# --- Gallery (Tabname) Logic ---
def refresh_gallery_list():
    """Refreshes Tabname gallery AND reloads SIFT templates."""
    try:
        for item in gallery_image_list.get_children():
            gallery_image_list.delete(item)
        files = os.listdir(TABNAME_DIR)
        png_files = sorted([f for f in files if f.endswith('.png')], reverse=True)
        for filename in png_files:
            gallery_image_list.insert("", tk.END, text=filename, iid=filename)
        load_all_sift_templates()
    except Exception as e:
        update_status('status_error', str(e))

def on_gallery_item_select(event):
    global gallery_preview_photo
    try:
        selected_items = gallery_image_list.selection()
        if not selected_items:
            gallery_preview_label.config(image='', text=translations['gallery_placeholder'][current_lang])
            gallery_rename_button.config(state=tk.DISABLED)
            gallery_delete_button.config(state=tk.DISABLED)
            return
        filename = selected_items[0]
        filepath = os.path.join(TABNAME_DIR, filename)
        img = Image.open(filepath)
        max_w = gallery_preview_label.winfo_width()
        max_h = gallery_preview_label.winfo_height()
        if max_w <= 1: max_w = 300
        if max_h <= 1: max_h = 300
        img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
        gallery_preview_photo = ImageTk.PhotoImage(img)
        gallery_preview_label.config(image=gallery_preview_photo, text="")
        gallery_preview_label.image = gallery_preview_photo
        gallery_rename_button.config(state=tk.NORMAL)
        gallery_delete_button.config(state=tk.NORMAL)
    except Exception as e:
        gallery_preview_label.config(image='', text=f"Error loading:\n{e}")
        gallery_rename_button.config(state=tk.DISABLED)
        gallery_delete_button.config(state=tk.DISABLED)

def rename_gallery_item():
    try:
        selected_items = gallery_image_list.selection()
        if not selected_items: return
        old_filename = selected_items[0]
        old_path = os.path.join(TABNAME_DIR, old_filename)
        dialog = AskROINameDialog(root,
            translations['rename_prompt_title'][current_lang],
            translations['rename_prompt_text'][current_lang],
            PREDEFINED_TABNAME_NAMES,
            initialvalue=old_filename.replace(".png", "")
        )
        new_filename = dialog.result
        if new_filename: new_filename = new_filename.strip()
            
        if new_filename and new_filename != old_filename.replace(".png", ""):
            if not new_filename.endswith('.png'): new_filename += '.png'
            new_path = os.path.join(TABNAME_DIR, new_filename)
            if os.path.exists(new_path):
                messagebox.showerror("Error", f"File '{new_filename}' already exists.")
                return
            os.rename(old_path, new_path)
            refresh_gallery_list()
            gallery_image_list.selection_set(new_filename)
    except Exception as e:
        update_status('status_error', str(e))

def delete_gallery_item():
    try:
        selected_items = gallery_image_list.selection()
        if not selected_items: return
        filename = selected_items[0]
        if not messagebox.askyesno(
            translations['delete_confirm_title'][current_lang],
            translations['delete_confirm_text'][current_lang].format(content=filename)
        ):
            return
        filepath = os.path.join(TABNAME_DIR, filename)
        os.remove(filepath)
        refresh_gallery_list()
        on_gallery_item_select(None)
    except Exception as e:
        update_status('status_error', str(e))

# --- ROI Set File Logic ---
def refresh_roi_file_list():
    """Refreshes ROI Set list ONLY."""
    try:
        for item in roi_set_list.get_children():
            roi_set_list.delete(item)
        files = os.listdir(ROI_DIR)
        json_files = sorted([f for f in files if f.endswith('.json')], reverse=True)
        for filename in json_files:
            roi_set_list.insert("", tk.END, text=filename, iid=filename)
        on_roi_set_select(None)
    except Exception as e:
        update_status('status_error', str(e))

def on_roi_set_select(event):
    """Enables/disables ROI action buttons based on selection."""
    selected_items = roi_set_list.selection()
    if selected_items:
        add_to_roi_set_button.config(state=tk.NORMAL)
        roi_rename_button.config(state=tk.NORMAL)
        roi_delete_button.config(state=tk.NORMAL)
    else:
        add_to_roi_set_button.config(state=tk.DISABLED)
        roi_rename_button.config(state=tk.DISABLED)
        roi_delete_button.config(state=tk.DISABLED)

def rename_roi_file():
    """Renames an ROI .json file."""
    try:
        selected_items = roi_set_list.selection()
        if not selected_items: return
        old_filename = selected_items[0]
        old_path = os.path.join(ROI_DIR, old_filename)
        dialog = AskROINameDialog(root,
            translations['rename_prompt_title'][current_lang],
            translations['rename_prompt_text'][current_lang],
            PREDEFINED_TABNAME_NAMES,
            initialvalue=old_filename.replace(".json", "")
        )
        new_filename = dialog.result
        if new_filename: new_filename = new_filename.strip()
            
        if new_filename and new_filename != old_filename.replace(".json", ""):
            if not new_filename.endswith('.json'): new_filename += '.json'
            new_path = os.path.join(ROI_DIR, new_filename)
            if os.path.exists(new_path):
                messagebox.showerror("Error", f"File '{new_filename}' already exists.")
                return
            os.rename(old_path, new_path)
            refresh_roi_file_list()
            roi_set_list.selection_set(new_filename)
    except Exception as e:
        update_status('status_error', str(e))

def delete_roi_file():
    """Deletes an ROI .json file."""
    try:
        selected_items = roi_set_list.selection()
        if not selected_items: return
        filename = selected_items[0]
        if not messagebox.askyesno(
            translations['delete_confirm_title'][current_lang],
            translations['delete_confirm_text'][current_lang].format(content=filename)
        ):
            return
        filepath = os.path.join(ROI_DIR, filename)
        os.remove(filepath)
        refresh_roi_file_list()
    except Exception as e:
        update_status('status_error', str(e))

# --- Status Template Management Logic ---
def refresh_status_folders():
    """Refreshes the folder list in the Status Tab."""
    try:
        for item in status_folder_list.get_children():
            status_folder_list.delete(item)
        folders = [f for f in os.listdir(STATUS_TEMPLATE_DIR) if os.path.isdir(os.path.join(STATUS_TEMPLATE_DIR, f))]
        folders.sort()
        for foldername in folders:
            status_folder_list.insert("", tk.END, text=foldername, iid=foldername)
        for item in status_image_list.get_children():
            status_image_list.delete(item)
        status_add_image_button.config(state=tk.DISABLED)
        status_rename_image_button.config(state=tk.DISABLED)
        status_delete_image_button.config(state=tk.DISABLED)
        status_rename_folder_button.config(state=tk.DISABLED)
        status_delete_folder_button.config(state=tk.DISABLED)
        status_preview_label.config(image='', text=translations['gallery_placeholder'][current_lang])
    except Exception as e:
        update_status('status_error', str(e))

def on_status_folder_select(event):
    """Populates the image list based on the selected folder."""
    global status_preview_photo
    status_preview_photo = None
    status_rename_image_button.config(state=tk.DISABLED)
    status_delete_image_button.config(state=tk.DISABLED)
    try:
        for item in status_image_list.get_children():
            status_image_list.delete(item)
        selected_items = status_folder_list.selection()
        if not selected_items:
            status_add_image_button.config(state=tk.DISABLED)
            status_rename_folder_button.config(state=tk.DISABLED)
            status_delete_folder_button.config(state=tk.DISABLED)
            return
        foldername = selected_items[0]
        folder_path = os.path.join(STATUS_TEMPLATE_DIR, foldername)
        files = os.listdir(folder_path)
        png_files = sorted([f for f in files if f.endswith('.png')])
        for filename in png_files:
            status_image_list.insert("", tk.END, text=filename, iid=filename)
        status_add_image_button.config(state=tk.NORMAL)
        status_rename_folder_button.config(state=tk.NORMAL)
        status_delete_folder_button.config(state=tk.NORMAL)
    except Exception as e:
        update_status('status_error', str(e))
        
def on_status_image_select(event):
    """Shows a preview of the selected status image."""
    global status_preview_photo
    try:
        folder_items = status_folder_list.selection()
        image_items = status_image_list.selection()
        if not folder_items or not image_items:
            status_rename_image_button.config(state=tk.DISABLED)
            status_delete_image_button.config(state=tk.DISABLED)
            return
        foldername = folder_items[0]
        filename = image_items[0]
        filepath = os.path.join(STATUS_TEMPLATE_DIR, foldername, filename)
        img = Image.open(filepath)
        max_w = status_preview_label.winfo_width()
        max_h = status_preview_label.winfo_height()
        if max_w <= 1: max_w = 200
        if max_h <= 1: max_h = 200
        img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
        status_preview_photo = ImageTk.PhotoImage(img)
        status_preview_label.config(image=status_preview_photo, text="")
        status_preview_label.image = status_preview_photo
        status_rename_image_button.config(state=tk.NORMAL)
        status_delete_image_button.config(state=tk.NORMAL)
    except Exception as e:
        status_preview_label.config(image='', text=f"Error loading:\n{e}")
        status_rename_image_button.config(state=tk.DISABLED)
        status_delete_image_button.config(state=tk.DISABLED)

def create_status_folder():
    """Creates a new folder in pictures/status/"""
    dialog = AskROINameDialog(root,
        translations['create_folder_button'][current_lang],
        translations['rename_prompt_text'][current_lang],
        PREDEFINED_TABNAME_NAMES
    )
    foldername = dialog.result
    if foldername: foldername = foldername.strip()
    if foldername:
        try:
            folder_path = os.path.join(STATUS_TEMPLATE_DIR, foldername)
            os.makedirs(folder_path, exist_ok=True)
            refresh_status_folders()
            status_folder_list.selection_set(foldername)
        except Exception as e:
            update_status('status_error', str(e))

def rename_status_folder():
    """Renames a folder in pictures/status/"""
    try:
        selected_items = status_folder_list.selection()
        if not selected_items: return
        old_foldername = selected_items[0]
        old_path = os.path.join(STATUS_TEMPLATE_DIR, old_foldername)
        dialog = AskROINameDialog(root,
            translations['rename_folder_button'][current_lang],
            translations['rename_prompt_text'][current_lang],
            PREDEFINED_TABNAME_NAMES,
            initialvalue=old_foldername
        )
        new_foldername = dialog.result
        if new_foldername: new_foldername = new_foldername.strip()
        if new_foldername and new_foldername != old_foldername:
            new_path = os.path.join(STATUS_TEMPLATE_DIR, new_foldername)
            if os.path.exists(new_path):
                messagebox.showerror("Error", f"Folder '{new_foldername}' already exists.")
                return
            os.rename(old_path, new_path)
            refresh_status_folders()
            status_folder_list.selection_set(new_foldername)
            load_all_sift_templates()
    except Exception as e:
        update_status('status_error', str(e))

def delete_status_folder():
    """Deletes a folder (and its contents) from pictures/status/"""
    try:
        selected_items = status_folder_list.selection()
        if not selected_items: return
        foldername = selected_items[0]
        if not messagebox.askyesno(
            translations['delete_confirm_title'][current_lang],
            translations['delete_folder_confirm_text'][current_lang].format(content=foldername)
        ):
            return
        folder_path = os.path.join(STATUS_TEMPLATE_DIR, foldername)
        shutil.rmtree(folder_path)
        refresh_status_folders()
        load_all_sift_templates()
    except Exception as e:
        update_status('status_error', str(e))

def start_add_status_image():
    """Captures a region and saves it as a new status image."""
    selected_folders = status_folder_list.selection()
    if not selected_folders:
        messagebox.showwarning("No Folder", translations['select_folder_prompt'][current_lang])
        return
    foldername = selected_folders[0]
    
    root.withdraw()
    time.sleep(0.5)
    selector = RegionSelector(root)
    root.wait_window(selector.selector_window)
    root.deiconify()
    
    if selector.box:
        try:
            cropped_image = selector.background_image.crop(selector.box)
            dialog = AskROINameDialog(root,
                translations['add_image_prompt_title'][current_lang],
                translations['add_image_prompt_text'][current_lang],
                PREDEFINED_STATUS_NAMES
            )
            imagename = dialog.result
            if imagename: imagename = imagename.strip()
                
            if not imagename:
                return
            if not imagename.endswith('.png'):
                imagename += '.png'
            save_path = os.path.join(STATUS_TEMPLATE_DIR, foldername, imagename)
            cropped_image.save(save_path)
            on_status_folder_select(None)
            status_image_list.selection_set(imagename)
            load_all_sift_templates() 
        except Exception as e:
            update_status('status_error', str(e))

def rename_status_image():
    """Renames a status image file."""
    try:
        folder_items = status_folder_list.selection()
        image_items = status_image_list.selection()
        if not folder_items or not image_items: return
        foldername = folder_items[0]
        old_filename = image_items[0]
        old_path = os.path.join(STATUS_TEMPLATE_DIR, foldername, old_filename)
        dialog = AskROINameDialog(root,
            translations['rename_image_button'][current_lang],
            translations['rename_prompt_text'][current_lang],
            PREDEFINED_STATUS_NAMES,
            initialvalue=old_filename.replace(".png", "")
        )
        new_filename = dialog.result
        if new_filename: new_filename = new_filename.strip()
            
        if new_filename and new_filename != old_filename.replace(".png", ""):
            if not new_filename.endswith('.png'): new_filename += '.png'
            new_path = os.path.join(STATUS_TEMPLATE_DIR, foldername, new_filename)
            if os.path.exists(new_path):
                messagebox.showerror("Error", f"File '{new_filename}' already exists.")
                return
            os.rename(old_path, new_path)
            on_status_folder_select(None)
            status_image_list.selection_set(new_filename)
            load_all_sift_templates()
    except Exception as e:
        update_status('status_error', str(e))

def delete_status_image():
    """Deletes a status image file."""
    try:
        folder_items = status_folder_list.selection()
        image_items = status_image_list.selection()
        if not folder_items or not image_items: return
        foldername = folder_items[0]
        filename = image_items[0]
        if not messagebox.askyesno(
            translations['delete_confirm_title'][current_lang],
            translations['delete_confirm_text'][current_lang].format(content=filename)
        ):
            return
        filepath = os.path.join(STATUS_TEMPLATE_DIR, foldername, filename)
        os.remove(filepath)
        on_status_folder_select(None)
        load_all_sift_templates()
    except Exception as e:
        update_status('status_error', str(e))

# --- OCR Debug Tab Logic ---
# --- (MODIFIED) - ล้าง 4 ช่อง (2x2 grid) ---
def clear_ocr_debug_tab():
    """(MODIFIED) Resets the OCR debug tab to its default state (4-grid)."""
    global ocr_raw_photo, ocr_gray_photo, ocr_contrast_photo, ocr_final_photo
    ocr_raw_photo = None
    ocr_gray_photo = None
    ocr_contrast_photo = None
    ocr_final_photo = None
    
    ocr_split_combo.set("")
    ocr_roi_combo.set("")
    ocr_roi_combo['values'] = []
    
    ocr_raw_image_label.config(image=None, text="...")
    ocr_gray_image_label.config(image=None, text="...")
    ocr_contrast_image_label.config(image=None, text="...")
    ocr_final_image_label.config(image=None, text="...")
    ocr_result_label.config(text="")

def refresh_ocr_debug_splits():
    """Loads the latest capture data into the Split dropdown."""
    if not g_latest_sift_results:
        messagebox.showinfo("No Data", translations['ocr_no_data'][current_lang])
        return
        
    split_names = []
    for i, (_, match_name, _, _, _) in enumerate(g_latest_sift_results):
        split_names.append(f"Split {i+1}: {match_name}")
        
    ocr_split_combo['values'] = split_names
    clear_ocr_debug_tab()
    ocr_split_combo.current(0)
    on_ocr_split_select(None)

def on_ocr_split_select(event):
    """Populates the ROI dropdown based on the selected split."""
    try:
        selected_index = ocr_split_combo.current()
        if selected_index == -1:
            clear_ocr_debug_tab()
            return
            
        _pil_image, match_name, _offset, data_results, _validation = g_latest_sift_results[selected_index]
        
        roi_filename = match_name.replace(".png", "") + ".json"
        roi_filepath = os.path.join(ROI_DIR, roi_filename)
        
        roi_keys = []
        if os.path.exists(roi_filepath):
            with open(roi_filepath, 'r', encoding='utf-8') as f:
                roi_file_data = json.load(f)
                roi_keys = list(roi_file_data.keys())
        
        ocr_roi_combo['values'] = roi_keys
        if roi_keys:
            ocr_roi_combo.current(0)
            on_ocr_roi_select(None)
        else:
            clear_ocr_debug_tab()
            
    except Exception as e:
        print(f"Error on split select: {e}")
        clear_ocr_debug_tab()

# --- (MODIFIED) - แสดงผล 4 ขั้นตอน (2x2 grid) ---
def on_ocr_roi_select(event):
    """(MODIFIED) Performs the full debug process and displays all 4 steps."""
    global ocr_raw_photo, ocr_gray_photo, ocr_contrast_photo, ocr_final_photo
    try:
        split_index = ocr_split_combo.current()
        roi_key = ocr_roi_combo.get()
        
        if split_index == -1 or not roi_key:
            return

        split_pil_image, match_name, (crop_offset_x, crop_offset_y), _, _ = g_latest_sift_results[split_index]

        roi_filename = match_name.replace(".png", "") + ".json"
        roi_filepath = os.path.join(ROI_DIR, roi_filename)
        
        if not os.path.exists(roi_filepath):
            raise FileNotFoundError(f"{roi_filename} not found")
            
        with open(roi_filepath, 'r', encoding='utf-8') as f:
            roi_file_data = json.load(f)
            
        if roi_key not in roi_file_data:
            raise KeyError(f"{roi_key} not in {roi_filename}")
            
        [global_x, global_y, global_w, global_h] = roi_file_data[roi_key]

        local_x = global_x - crop_offset_x
        local_y = global_y - crop_offset_y
        roi_box_pil = (local_x, local_y, local_x + global_w, local_y + global_h)
        
        # --- (NEW) Call the preprocessing function ---
        # (MODIFIED) Pass roi_key to preprocess_for_ocr
        processing_steps = preprocess_for_ocr(split_pil_image.crop(roi_box_pil), roi_key) 
        if processing_steps is None:
            raise ValueError("Preprocessing failed")
            
        # Get all intermediate images
        raw_pil = processing_steps['raw_pil']
        gray_pil = Image.fromarray(processing_steps['gray'])
        contrast_pil = Image.fromarray(processing_steps['contrast'])
        final_pil = Image.fromarray(processing_steps['final'])
        final_for_ocr_cv = processing_steps['final'] # This is sent to OCR
            
        # Get OCR result from the final image
        ocr_results = ocr_reader.readtext(final_for_ocr_cv, allowlist=OCR_ALLOWLIST, detail=0)
        extracted_text = "".join(ocr_results).strip()
        if not extracted_text:
            extracted_text = "N/A"

        # Create PhotoImage objects for the 4-grid display
        img_size = (root.winfo_width() // 2 - 50, root.winfo_height() // 3) # (Dynamic resize 2-wide)
        if img_size[0] < 100: img_size = (200, 100) # Minimum size
        if img_size[1] < 50: img_size = (img_size[0], 100)
            
        resample_method = Image.Resampling.NEAREST
        
        ocr_raw_photo = ImageTk.PhotoImage(raw_pil.resize(img_size, resample_method))
        ocr_gray_photo = ImageTk.PhotoImage(gray_pil.resize(img_size, resample_method))
        ocr_contrast_photo = ImageTk.PhotoImage(contrast_pil.resize(img_size, resample_method))
        ocr_final_photo = ImageTk.PhotoImage(final_pil.resize(img_size, resample_method))
        
        # Update all 4 image labels
        ocr_raw_image_label.config(image=ocr_raw_photo, text="")
        ocr_raw_image_label.image = ocr_raw_photo
        ocr_gray_image_label.config(image=ocr_gray_photo, text="")
        ocr_gray_image_label.image = ocr_gray_photo
        ocr_contrast_image_label.config(image=ocr_contrast_photo, text="")
        ocr_contrast_image_label.image = ocr_contrast_photo
        ocr_final_image_label.config(image=ocr_final_photo, text="")
        ocr_final_image_label.image = ocr_final_photo
        
        ocr_result_label.config(text=extracted_text)
        
    except Exception as e:
        print(f"Error on ROI select: {e}")
        ocr_result_label.config(text=f"Error: {e}")
        ocr_raw_image_label.config(image=None, text="Error")
        ocr_gray_image_label.config(image=None, text="Error")
        ocr_contrast_image_label.config(image=None, text="Error")
        ocr_final_image_label.config(image=None, text="Error")

# --- General Functions ---
def update_status(text_key, dynamic_content=""):
    try:
        base_text = translations.get(text_key, {}).get(current_lang, text_key)
        final_text = base_text.format(content=dynamic_content)
        status_label.config(text=final_text)
    except tk.TclError: pass

def on_closing():
    if is_running:
        if messagebox.askyesno(
            translations['confirm_close_title'][current_lang], 
            translations['confirm_close_message'][current_lang]
        ):
            stop_capture()
            root.destroy()
    else:
        root.destroy()

# ---- 1. สร้างหน้าต่างหลัก และ Style ----
root = tk.Tk()
root.geometry("1000x850") # (MODIFIED) Make window larger for new settings UI
root.resizable(True, True) 
font_family = "Segoe UI" if sys.platform == "win32" else "Arial"
style = ttk.Style()
style.theme_use('clam') 
style.configure('.', font=(font_family, 10))
style.configure('TFrame', background='#f0f0f0')
style.configure('TLabel', background='#f0f0f0')
style.configure('TButton', font=(font_family, 10, 'bold'), padding=5)
style.configure('TEntry', font=(font_family, 10))
style.configure('Horizontal.TProgressbar', thickness=15)
style.configure('TCombobox', font=(font_family, 10))
style.configure('TNotebook.Tab', font=(font_family, 10, 'bold'), padding=[10, 5])
style.configure('Treeview.Heading', font=(font_family, 11, 'bold'))
style.configure('Bold.TLabel', font=(font_family, 11, 'bold'))
style.configure('Data.TLabel', font=(font_family, 9))
style.configure('Status.TLabel', font=(font_family, 10, 'bold'))
style.configure('Help.TLabel', font=(font_family, 9), foreground='#555555') # (NEW)
style.configure('TListbox', font=(font_family, 10))

# ---- 2. สร้าง Notebook (Tabbed Interface) ----
notebook = ttk.Notebook(root, padding=10)
notebook.pack(fill=tk.BOTH, expand=True)

# ---- 3. สร้าง Tab 1: Auto-Capture ----
capture_tab = ttk.Frame(notebook)
notebook.add(capture_tab, text="Capture")
settings_frame = ttk.Frame(capture_tab, padding=(0, 0, 0, 10))
settings_frame.pack(fill=tk.X)
interval_label = ttk.Label(settings_frame)
interval_label.pack(side=tk.LEFT, padx=(0, 5))
interval_entry = EntryWithRightClickMenu(settings_frame, width=5, font=(font_family, 10))
interval_entry.pack(side=tk.LEFT, padx=5)
interval_entry.insert(0, "5") 
start_button = ttk.Button(settings_frame, command=start_capture)
start_button.pack(side=tk.LEFT, padx=5)
stop_button = ttk.Button(settings_frame, command=stop_capture, state=tk.DISABLED)
stop_button.pack(side=tk.LEFT, padx=5)
capture_region_button = ttk.Button(settings_frame, command=start_region_capture)
capture_region_button.pack(side=tk.LEFT, padx=15)

minimize_on_start_var = tk.BooleanVar()
minimize_on_start_check = ttk.Checkbutton(settings_frame, variable=minimize_on_start_var)
minimize_on_start_check.pack(side=tk.LEFT, padx=10)

lang_button = ttk.Button(settings_frame, command=toggle_language)
lang_button.pack(side=tk.RIGHT, padx=5) 
split_frame = ttk.Frame(capture_tab, padding=(0, 8, 0, 0))
split_frame.pack(fill=tk.X)
split_method_label = ttk.Label(split_frame)
split_method_label.pack(side=tk.LEFT, padx=(0, 5))
split_method_combo = ttk.Combobox(split_frame, state="readonly", font=(font_family, 10))
split_method_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
progress_frame = ttk.Frame(capture_tab, padding=(0, 10, 0, 5))
progress_frame.pack(fill=tk.X)
progress_label = ttk.Label(progress_frame)
progress_label.pack(side=tk.LEFT, padx=(0, 5))
progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode='determinate')
progress_bar.pack(fill=tk.X, expand=True)

# (MODIFIED) - ใช้ ScrollableFrame เพื่อให้เลื่อนดูข้อมูลแนวตั้งได้ด้วย
crop_display_frame_scroller = ScrollableFrame(capture_tab)
crop_display_frame_scroller.pack(fill=tk.BOTH, expand=True, pady=5)
crop_display_frame = crop_display_frame_scroller.interior # นี่คือ frame จริงที่เราจะ pack รูปใส่

# ---- 4. สร้าง Tab 2: Tabname (Gallery) ----
gallery_tab = ttk.Frame(notebook)
notebook.add(gallery_tab, text="Tabname")
gallery_paned_window = ttk.PanedWindow(gallery_tab, orient=tk.HORIZONTAL)
gallery_paned_window.pack(fill=tk.BOTH, expand=True)
gallery_list_frame = ttk.Frame(gallery_paned_window, padding=5)
gallery_paned_window.add(gallery_list_frame, weight=1)
gallery_refresh_button = ttk.Button(gallery_list_frame, command=refresh_gallery_list)
gallery_refresh_button.pack(fill=tk.X, pady=5)
gallery_image_list = ttk.Treeview(gallery_list_frame, selectmode="browse")
gallery_image_list.heading("#0", text="")
gallery_image_list.bind("<<TreeviewSelect>>", on_gallery_item_select)
# (NEW) Scrollbar for Gallery
gallery_scrollbar = ttk.Scrollbar(gallery_list_frame, orient=tk.VERTICAL, command=gallery_image_list.yview)
gallery_image_list.configure(yscrollcommand=gallery_scrollbar.set)
gallery_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
gallery_image_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

gallery_preview_frame = ttk.Frame(gallery_paned_window, padding=10)
gallery_paned_window.add(gallery_preview_frame, weight=2)
gallery_preview_header_label = ttk.Label(gallery_preview_frame, font=(font_family, 12, 'bold'))
gallery_preview_header_label.pack(pady=5)
gallery_preview_label = tk.Label(gallery_preview_frame, background="#ffffff", relief=tk.SUNKEN, borderwidth=1, width=40, height=20)
gallery_preview_label.pack(fill=tk.BOTH, expand=True)
gallery_action_frame = ttk.Frame(gallery_preview_frame, padding=(0, 10, 0, 0))
gallery_action_frame.pack(fill=tk.X)
gallery_rename_button = ttk.Button(gallery_action_frame, command=rename_gallery_item, state=tk.DISABLED)
gallery_rename_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
gallery_delete_button = ttk.Button(gallery_action_frame, command=delete_gallery_item, state=tk.DISABLED)
gallery_delete_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

# ---- 5. สร้าง Tab 3: ROI Sets ----
roi_tab = ttk.Frame(notebook)
notebook.add(roi_tab, text="ROI Sets")
roi_action_frame = ttk.Frame(roi_tab, padding=5)
roi_action_frame.pack(fill=tk.X)
create_roi_set_button = ttk.Button(roi_action_frame, command=start_roi_set_creation)
create_roi_set_button.pack(side=tk.LEFT, padx=5, pady=5)
add_to_roi_set_button = ttk.Button(roi_action_frame, command=start_add_to_roi_set, state=tk.DISABLED)
add_to_roi_set_button.pack(side=tk.LEFT, padx=5, pady=5)
roi_refresh_button = ttk.Button(roi_action_frame, command=refresh_roi_file_list)
roi_refresh_button.pack(side=tk.LEFT, padx=5, pady=5)
roi_paned_window = ttk.PanedWindow(roi_tab, orient=tk.HORIZONTAL)
roi_paned_window.pack(fill=tk.BOTH, expand=True)
roi_list_frame = ttk.Frame(roi_paned_window, padding=5)
roi_paned_window.add(roi_list_frame, weight=1)
roi_set_list = ttk.Treeview(roi_list_frame, selectmode="browse")
roi_set_list.heading("#0", text="ROI Set Files")
roi_set_list.bind("<<TreeviewSelect>>", on_roi_set_select)
# (NEW) Scrollbar for ROI
roi_scrollbar = ttk.Scrollbar(roi_list_frame, orient=tk.VERTICAL, command=roi_set_list.yview)
roi_set_list.configure(yscrollcommand=roi_scrollbar.set)
roi_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
roi_set_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

roi_edit_frame = ttk.Frame(roi_paned_window, padding=10)
roi_paned_window.add(roi_edit_frame, weight=1)
ttk.Label(roi_edit_frame, text="Selected File Actions:", font=(font_family, 11, 'bold')).pack(pady=10)
roi_rename_button = ttk.Button(roi_edit_frame, command=rename_roi_file, state=tk.DISABLED)
roi_rename_button.pack(fill=tk.X, padx=5, pady=5)
roi_delete_button = ttk.Button(roi_edit_frame, command=delete_roi_file, state=tk.DISABLED)
roi_delete_button.pack(fill=tk.X, padx=5, pady=5)

# ---- 6. สร้าง Tab 4: Status Templates ----
status_tab = ttk.Frame(notebook)
notebook.add(status_tab, text="Status Templates")
status_paned_window = ttk.PanedWindow(status_tab, orient=tk.HORIZONTAL)
status_paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
status_folder_frame = ttk.Frame(status_paned_window, padding=5)
status_paned_window.add(status_folder_frame, weight=1)
status_folder_actions = ttk.Frame(status_folder_frame)
status_folder_actions.pack(fill=tk.X, pady=2)
status_create_folder_button = ttk.Button(status_folder_actions, command=create_status_folder)
status_create_folder_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
status_rename_folder_button = ttk.Button(status_folder_actions, command=rename_status_folder, state=tk.DISABLED)
status_rename_folder_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
status_delete_folder_button = ttk.Button(status_folder_actions, command=delete_status_folder, state=tk.DISABLED)
status_delete_folder_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
status_folder_list = ttk.Treeview(status_folder_frame, selectmode="browse")
status_folder_list.heading("#0", text="Tabname Folders")
status_folder_list.bind("<<TreeviewSelect>>", on_status_folder_select)
# (NEW) Scrollbar for Status Folder
status_folder_scrollbar = ttk.Scrollbar(status_folder_frame, orient=tk.VERTICAL, command=status_folder_list.yview)
status_folder_list.configure(yscrollcommand=status_folder_scrollbar.set)
status_folder_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
status_folder_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)

status_image_frame = ttk.Frame(status_paned_window, padding=5)
status_paned_window.add(status_image_frame, weight=2)
status_image_actions = ttk.Frame(status_image_frame)
status_image_actions.pack(fill=tk.X, pady=2)
status_add_image_button = ttk.Button(status_image_actions, command=start_add_status_image, state=tk.DISABLED)
status_add_image_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
status_rename_image_button = ttk.Button(status_image_actions, command=rename_status_image, state=tk.DISABLED)
status_rename_image_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
status_delete_image_button = ttk.Button(status_image_actions, command=delete_status_image, state=tk.DISABLED)
status_delete_image_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
status_image_list = ttk.Treeview(status_image_frame, selectmode="browse")
status_image_list.heading("#0", text="Status Images")
status_image_list.bind("<<TreeviewSelect>>", on_status_image_select)
# (NEW) Scrollbar for Status Image
status_image_scrollbar = ttk.Scrollbar(status_image_frame, orient=tk.VERTICAL, command=status_image_list.yview)
status_image_list.configure(yscrollcommand=status_image_scrollbar.set)
status_image_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
status_image_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)

status_preview_label = tk.Label(status_image_frame, background="#ffffff", relief=tk.SUNKEN, borderwidth=1)
status_preview_label.pack(fill=tk.BOTH, expand=True, pady=5)

# ---- 7. สร้าง Tab 5: Settings (MODIFIED) ----
settings_tab_scroller = ScrollableFrame(notebook) # (MODIFIED) ใช้ ScrollableFrame
settings_tab = settings_tab_scroller.interior # (NEW)
settings_tab.config(padding=10) # (NEW)
notebook.add(settings_tab_scroller, text="Settings") # (MODIFIED)

# --- SIFT Frame ---
sift_frame = ttk.Frame(settings_tab)
sift_frame.pack(fill=tk.X, pady=10)
ttk.Label(sift_frame, text="SIFT Matching Settings", style='Bold.TLabel').pack(anchor=tk.W, pady=(0, 5))
tabname_threshold_label = ttk.Label(sift_frame, text="Tabname SIFT Threshold:", anchor=tk.W)
tabname_threshold_label.pack(fill=tk.X)
tabname_threshold_entry = EntryWithRightClickMenu(sift_frame, width=10)
tabname_threshold_entry.pack(anchor=tk.W, pady=(5, 10))
status_threshold_label = ttk.Label(sift_frame, text="Status SIFT Threshold:", anchor=tk.W)
status_threshold_label.pack(fill=tk.X)
status_threshold_entry = EntryWithRightClickMenu(sift_frame, width=10)
status_threshold_entry.pack(anchor=tk.W, pady=(5, 10))

ttk.Separator(settings_tab, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

# --- (MODIFIED) OCR Settings Frame (Part 1: Basic) ---
ocr_settings_frame = ttk.Frame(settings_tab)
ocr_settings_frame.pack(fill=tk.X, pady=10)
ocr_settings_header = ttk.Label(ocr_settings_frame, style='Bold.TLabel')
ocr_settings_header.pack(anchor=tk.W, pady=(0, 10))

# 1. Scale
ocr_scale_label = ttk.Label(ocr_settings_frame, anchor=tk.W)
ocr_scale_label.pack(fill=tk.X)
ocr_scale_entry = EntryWithRightClickMenu(ocr_settings_frame, width=10)
ocr_scale_entry.pack(anchor=tk.W, pady=2)
ocr_scale_help = ttk.Label(ocr_settings_frame, style='Help.TLabel', anchor=tk.W)
ocr_scale_help.pack(fill=tk.X, pady=(0, 10))

# 2. CLAHE
ocr_clahe_label = ttk.Label(ocr_settings_frame, anchor=tk.W)
ocr_clahe_label.pack(fill=tk.X)
ocr_clahe_entry = EntryWithRightClickMenu(ocr_settings_frame, width=10)
ocr_clahe_entry.pack(anchor=tk.W, pady=2)
ocr_clahe_help = ttk.Label(ocr_settings_frame, style='Help.TLabel', anchor=tk.W)
ocr_clahe_help.pack(fill=tk.X, pady=(0, 10))

# 3. Median
ocr_median_label = ttk.Label(ocr_settings_frame, anchor=tk.W)
ocr_median_label.pack(fill=tk.X)
ocr_median_entry = EntryWithRightClickMenu(ocr_settings_frame, width=10)
ocr_median_entry.pack(anchor=tk.W, pady=2)
ocr_median_help = ttk.Label(ocr_settings_frame, style='Help.TLabel', anchor=tk.W)
ocr_median_help.pack(fill=tk.X, pady=(0, 10))

# 4. Opening
ocr_opening_label = ttk.Label(ocr_settings_frame, anchor=tk.W)
ocr_opening_label.pack(fill=tk.X)
ocr_opening_entry = EntryWithRightClickMenu(ocr_settings_frame, width=10)
ocr_opening_entry.pack(anchor=tk.W, pady=2)
ocr_opening_help = ttk.Label(ocr_settings_frame, style='Help.TLabel', anchor=tk.W)
ocr_opening_help.pack(fill=tk.X, pady=(0, 10))

# --- (NEW) OCR Settings Frame (Part 2: Kernel Sizes) ---
ocr_kernel_settings_header = ttk.Label(ocr_settings_frame, style='Bold.TLabel')
ocr_kernel_settings_header.pack(anchor=tk.W, pady=(10, 5))

# 1. Dilate/Thicken Kernel Size
ocr_dilate_label = ttk.Label(ocr_settings_frame, anchor=tk.W)
ocr_dilate_label.pack(fill=tk.X)
ocr_dilate_entry = EntryWithRightClickMenu(ocr_settings_frame, width=10)
ocr_dilate_entry.pack(anchor=tk.W, pady=2)
ocr_dilate_help = ttk.Label(ocr_settings_frame, style='Help.TLabel', anchor=tk.W)
ocr_dilate_help.pack(fill=tk.X, pady=(0, 10))

# 2. Erode/Thin Kernel Size
ocr_erode_label = ttk.Label(ocr_settings_frame, anchor=tk.W)
ocr_erode_label.pack(fill=tk.X)
ocr_erode_entry = EntryWithRightClickMenu(ocr_settings_frame, width=10)
ocr_erode_entry.pack(anchor=tk.W, pady=2)
ocr_erode_help = ttk.Label(ocr_settings_frame, style='Help.TLabel', anchor=tk.W)
ocr_erode_help.pack(fill=tk.X, pady=(0, 10))


ttk.Separator(settings_tab, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

# --- (NEW) OCR Settings Frame (Part 3: Target Selection) ---
ocr_targets_header = ttk.Label(settings_tab, style='Bold.TLabel')
ocr_targets_header.pack(anchor=tk.W, pady=(0, 10))

target_selection_frame = ttk.Frame(settings_tab)
target_selection_frame.pack(fill=tk.X)

# A. Available ROIs List
available_roi_frame = ttk.Frame(target_selection_frame)
available_roi_frame.pack(side=tk.LEFT, padx=5, fill=tk.BOTH, expand=True)
available_roi_label = ttk.Label(available_roi_frame)
available_roi_label.pack(fill=tk.X, pady=5)
available_roi_listbox = tk.Listbox(available_roi_frame, selectmode=tk.EXTENDED, height=15)
available_roi_listbox.pack(fill=tk.BOTH, expand=True)

# B. Control Buttons
control_button_frame = ttk.Frame(target_selection_frame, width=150)
control_button_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y)
control_button_frame.pack_propagate(False)

ttk.Label(control_button_frame, text="").pack(pady=10) # Spacer
add_dilate_button = ttk.Button(control_button_frame, command=add_dilate_target)
add_dilate_button.pack(fill=tk.X, pady=5)
remove_dilate_button = ttk.Button(control_button_frame, command=remove_dilate_target)
remove_dilate_button.pack(fill=tk.X, pady=5)

ttk.Label(control_button_frame, text="").pack(pady=10) # Spacer
add_erode_button = ttk.Button(control_button_frame, command=add_erode_target)
add_erode_button.pack(fill=tk.X, pady=5)
remove_erode_button = ttk.Button(control_button_frame, command=remove_erode_target)
remove_erode_button.pack(fill=tk.X, pady=5)

ttk.Label(control_button_frame, text="").pack(pady=10) # Spacer
refresh_targets_button = ttk.Button(control_button_frame, command=refresh_ocr_target_listboxes)
refresh_targets_button.pack(fill=tk.X, pady=5)

# C. Target Lists (Dilate and Erode)
target_list_frame = ttk.Frame(target_selection_frame)
target_list_frame.pack(side=tk.LEFT, padx=5, fill=tk.BOTH, expand=True)

# Dilate Targets
dilate_targets_label = ttk.Label(target_list_frame)
dilate_targets_label.pack(fill=tk.X, pady=5)
dilate_target_listbox = tk.Listbox(target_list_frame, selectmode=tk.EXTENDED, height=7)
dilate_target_listbox.pack(fill=tk.X, pady=(0, 5))

# Erode Targets
erode_targets_label = ttk.Label(target_list_frame)
erode_targets_label.pack(fill=tk.X, pady=5)
erode_target_listbox = tk.Listbox(target_list_frame, selectmode=tk.EXTENDED, height=7)
erode_target_listbox.pack(fill=tk.X)


ttk.Separator(settings_tab, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

# --- Google Sheet Frame ---
g_sheet_frame = ttk.Frame(settings_tab)
g_sheet_frame.pack(fill=tk.X, pady=10)
ttk.Label(g_sheet_frame, text="Google Sheet Export", style='Bold.TLabel').pack(anchor=tk.W, pady=(0, 5))
g_sheet_url_label = ttk.Label(g_sheet_frame, text="Google Sheet Web App URL:", anchor=tk.W)
g_sheet_url_label.pack(fill=tk.X)
g_sheet_url_entry = EntryWithRightClickMenu(g_sheet_frame, width=80)
g_sheet_url_entry.pack(fill=tk.X, pady=(5, 10))

# --- Save Button ---
g_sheet_save_button = ttk.Button(settings_tab, command=save_config)
g_sheet_save_button.pack(anchor=tk.W)


# ---- 8. สร้าง Tab 6: OCR Debug ----
ocr_debug_tab_scroller = ScrollableFrame(notebook) 
ocr_debug_tab = ocr_debug_tab_scroller.interior 
ocr_debug_tab.config(padding=10) 
notebook.add(ocr_debug_tab_scroller, text="OCR Debug") 

ocr_controls_frame = ttk.Frame(ocr_debug_tab)
ocr_controls_frame.pack(fill=tk.X, pady=5)
ocr_refresh_btn = ttk.Button(ocr_controls_frame, command=refresh_ocr_debug_splits)
ocr_refresh_btn.pack(side=tk.LEFT, padx=(0, 20))
ocr_split_label = ttk.Label(ocr_controls_frame, text="Select Split:")
ocr_split_label.pack(side=tk.LEFT, padx=(0, 5))
ocr_split_combo = ttk.Combobox(ocr_controls_frame, state="readonly", width=25)
ocr_split_combo.pack(side=tk.LEFT, padx=5)
ocr_roi_label = ttk.Label(ocr_controls_frame, text="Select ROI:")
ocr_roi_label.pack(side=tk.LEFT, padx=(10, 5))
ocr_roi_combo = ttk.Combobox(ocr_controls_frame, state="readonly", width=30)
ocr_roi_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
ocr_split_combo.bind("<<ComboboxSelected>>", on_ocr_split_select)
ocr_roi_combo.bind("<<ComboboxSelected>>", on_ocr_roi_select)

# --- (MODIFIED) สร้าง UI แบบ 2x2 Grid ---
ocr_top_row_frame = ttk.Frame(ocr_debug_tab)
ocr_top_row_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

ocr_raw_frame = ttk.Frame(ocr_top_row_frame, padding=5)
ocr_raw_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)
ocr_raw_frame_label = ttk.Label(ocr_raw_frame, text="Raw ROI", font=(font_family, 11, 'bold'))
ocr_raw_frame_label.pack(pady=5)
ocr_raw_image_label = tk.Label(ocr_raw_frame, background="#ffffff", relief=tk.SUNKEN, borderwidth=1)
ocr_raw_image_label.pack(fill=tk.BOTH, expand=True)

ocr_gray_frame = ttk.Frame(ocr_top_row_frame, padding=5) 
ocr_gray_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2) 
ocr_gray_frame_label = ttk.Label(ocr_gray_frame, text="Grayscale", font=(font_family, 11, 'bold')) 
ocr_gray_frame_label.pack(pady=5) 
ocr_gray_image_label = tk.Label(ocr_gray_frame, background="#ffffff", relief=tk.SUNKEN, borderwidth=1) 
ocr_gray_image_label.pack(fill=tk.BOTH, expand=True) 

ocr_bottom_row_frame = ttk.Frame(ocr_debug_tab)
ocr_bottom_row_frame.pack(fill=tk.BOTH, expand=True, pady=2)

ocr_contrast_frame = ttk.Frame(ocr_bottom_row_frame, padding=5) 
ocr_contrast_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2) 
ocr_contrast_frame_label = ttk.Label(ocr_contrast_frame, text="Conditional Processing", font=(font_family, 11, 'bold')) # (MODIFIED)
ocr_contrast_frame_label.pack(pady=5) 
ocr_contrast_image_label = tk.Label(ocr_contrast_frame, background="#ffffff", relief=tk.SUNKEN, borderwidth=1) 
ocr_contrast_image_label.pack(fill=tk.BOTH, expand=True) 

ocr_final_frame = ttk.Frame(ocr_bottom_row_frame, padding=5)
ocr_final_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)
ocr_final_frame_label = ttk.Label(ocr_final_frame, text="Final (Threshold)", font=(font_family, 11, 'bold'))
ocr_final_frame_label.pack(pady=5)
ocr_final_image_label = tk.Label(ocr_final_frame, background="#000000", relief=tk.SUNKEN, borderwidth=1)
ocr_final_image_label.pack(fill=tk.BOTH, expand=True)
# --- (END) สิ้นสุดการแก้ไข 2x2 Grid ---

ocr_result_frame = ttk.Frame(ocr_debug_tab)
ocr_result_frame.pack(fill=tk.X, pady=10)
ocr_result_text_label = ttk.Label(ocr_result_frame, text="OCR Result:", font=(font_family, 11, 'bold'))
ocr_result_text_label.pack(side=tk.LEFT)
ocr_result_label = ttk.Label(ocr_result_frame, text="", font=(font_family, 14, 'bold'), foreground="blue")
ocr_result_label.pack(side=tk.LEFT, padx=10)

# ---- 9. สร้างแถบสถานะ (ล่างสุด) ----
status_label = ttk.Label(root, relief=tk.SUNKEN, anchor=tk.W, padding=5, font=(font_family, 9))
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# ---- 10. ตั้งค่าการปิดหน้าต่าง และเริ่มแอป ----
root.protocol("WM_DELETE_WINDOW", on_closing) 
set_language(current_lang)
clear_image_display() 
refresh_gallery_list() # This loads ALL SIFT caches
refresh_roi_file_list()
refresh_status_folders()
on_gallery_item_select(None)
on_roi_set_select(None)
load_config() # Load all saved settings
root.mainloop()