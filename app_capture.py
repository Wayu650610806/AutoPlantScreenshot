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

# --- SIFT Global Initialization ---
try:
    sift = cv2.SIFT_create()
    tabname_sift_cache = {}
    status_sift_caches = {}
    TABNAME_SIFT_THRESHOLD = 70
    STATUS_SIFT_THRESHOLD = 15
except Exception as e:
    messagebox.showerror("OpenCV Error", f"ไม่สามารถเริ่ม SIFT ได้ (อาจต้องติดตั้ง opencv-contrib-python)\n{e}")
    sys.exit()

# --- EasyOCR Initialization ---
try:
    print("Loading EasyOCR Reader... (This may take a moment on first run)")
    ocr_reader = easyocr.Reader(['en']) 
    print("EasyOCR Reader loaded.")
except Exception as e:
    messagebox.showerror("EasyOCR Error", f"Could not initialize EasyOCR.\n{e}")
    sys.exit()
OCR_ALLOWLIST = '-.0123456789'

# --- Base Path Logic ---
def get_base_path():
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return base_path

# --- Global Paths & Constants ---
BASE_PATH = get_base_path()
TABNAME_DIR = os.path.join(BASE_PATH, "pictures", "tabname") 
STATUS_TEMPLATE_DIR = os.path.join(BASE_PATH, "pictures", "status")
ROI_DIR = os.path.join(BASE_PATH, "rois")
CONFIG_FILE_PATH = os.path.join(BASE_PATH, "config.json")
os.makedirs(TABNAME_DIR, exist_ok=True)
os.makedirs(STATUS_TEMPLATE_DIR, exist_ok=True)
os.makedirs(ROI_DIR, exist_ok=True)

# --- Predefined ROI Names ---
PREDEFINED_ROI_NAMES = [
    "乾溜ガス化炉A_温度_℃", "乾溜ガス化炉B_温度_℃", "乾溜ガス化炉C_温度_℃",
    "乾溜空気弁A_開度_%", "乾溜空気弁B_開度_%", "乾溜空気弁C_開度_%",
    "乾溜ガス化炉A_運転状況", "乾溜ガス化炉B_運転状況", "乾溜ガス化炉C_運転状況",
    "燃焼炉_温度_℃", "排ガス濃度_CO濃度_ppm"
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
    'app_title': {'en': 'Capture Tool v0.15 (Paste Fix)', 'ja': 'キャプチャーツール v0.15 (ペースト修正)'}, # (MODIFIED)
    'tab_capture': {'en': 'Auto-Capture', 'ja': '自動キャプチャ'},
    'tab_gallery': {'en': 'Tabname', 'ja': 'タブ名'}, 
    'tab_roi_sets': {'en': 'ROI Sets', 'ja': 'ROIセット'},
    'tab_status': {'en': 'Status Templates', 'ja': 'ステータス・テンプレート'},
    'tab_settings': {'en': 'Settings', 'ja': '設定'},
    'interval_label': {'en': 'Interval (sec):', 'ja': '間隔 (秒):'},
    'start_button': {'en': 'Start Auto', 'ja': '自動開始'},
    'stop_button': {'en': 'Stop Auto', 'ja': '自動停止'},
    'progress_label': {'en': 'Next capture in:', 'ja': '次のキャプチャ:'},
    'image_placeholder': {'en': 'Captured crops will appear here\n(Press "Start" to begin)', 'ja': 'キャプチャした画像はここに表示されます\n(「開始」を押してください)'},
    'split_method_label': {'en': 'Split Method:', 'ja': '分割方法:'},
    'capture_region_button': {'en': 'Capture Tabname', 'ja': 'タブ名キャプチャ'},
    'gallery_preview_header': {'en': 'Preview', 'ja': 'プレビュー'},
    'gallery_refresh_button': {'en': 'Refresh', 'ja': '更新'},
    'gallery_rename_button': {'en': 'Rename', 'ja': '名前変更'},
    'gallery_delete_button': {'en': 'Delete', 'ja': '削除'},
    'gallery_placeholder': {'en': 'Select an image to preview', 'ja': '画像を選択してください'},
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
    'g_sheet_url_label': {'en': 'Google Sheet Web App URL:', 'ja': 'Google Sheet Web AppのURL:'},
    'g_sheet_save_button': {'en': 'Save URL', 'ja': 'URL保存'},
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
    'confirm_close_title': {'en': 'Confirm Exit', 'ja': '終了確認'},
    'confirm_close_message': {'en': 'Auto-Capture is running. Are you sure you want to stop and exit?', 'ja': '自動キャプチャが実行中です。停止して終了しますか？'},
    'lang_button': {'en': '日本語', 'ja': 'English'},
    **{key: value for key, value in SPLIT_OPTIONS.items()} 
}

# --- Global Variables ---
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

# --- (NEW) Custom Entry with Right-Click ---
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

    def do_cut(self):
        self.event_generate("<<Cut>>")

    def do_copy(self):
        self.event_generate("<<Copy>>")

    def do_paste(self):
        self.event_generate("<<Paste>>")

# --- Custom Dialog for ROI Naming ---
class AskROINameDialog(simpledialog.Dialog):
    def __init__(self, parent, title, text, predefined_names):
        self.text = text
        self.predefined_names = predefined_names
        super().__init__(parent, title)
    def body(self, master):
        ttk.Label(master, text=self.text).pack(pady=5)
        # (MODIFIED) Use the new EntryWithRightClickMenu
        self.combo = ttk.Combobox(master, values=self.predefined_names, width=50)
        self.combo.pack(padx=10, pady=5)
        return self.combo
    def apply(self):
        self.result = self.combo.get()

# --- Config Persistence ---
def load_config():
    """Loads Google Sheet URL from config.json."""
    global g_sheet_url
    try:
        if os.path.exists(CONFIG_FILE_PATH):
            with open(CONFIG_FILE_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
                g_sheet_url = data.get("g_sheet_url", "")
                if g_sheet_url_entry:
                    g_sheet_url_entry.delete(0, tk.END)
                    g_sheet_url_entry.insert(0, g_sheet_url)
    except Exception as e:
        print(f"Error loading config: {e}")
        g_sheet_url = ""

def save_config():
    """Saves Google Sheet URL to config.json."""
    global g_sheet_url
    try:
        g_sheet_url = g_sheet_url_entry.get()
        data = {"g_sheet_url": g_sheet_url}
        with open(CONFIG_FILE_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
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
    notebook.tab(settings_tab, text=translations['tab_settings'][current_lang])
    
    # Capture Tab
    interval_label.config(text=translations['interval_label'][current_lang])
    start_button.config(text=translations['start_button'][current_lang])
    stop_button.config(text=translations['stop_button'][current_lang])
    lang_button.config(text=translations['lang_button'][current_lang])
    progress_label.config(text=translations['progress_label'][current_lang])
    split_method_label.config(text=translations['split_method_label'][current_lang])
    capture_region_button.config(text=translations['capture_region_button'][current_lang])
    
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

    # Settings Tab
    g_sheet_url_label.config(text=translations['g_sheet_url_label'][current_lang])
    g_sheet_save_button.config(text=translations['g_sheet_save_button'][current_lang])
    
    if not is_running:
        status_label.config(text=translations['status_idle'][current_lang])
        if image_placeholder_label:
            image_placeholder_label.config(text=translations['image_placeholder'][current_lang])
    else:
        interval = interval_entry.get()
        update_status('status_running', interval)

def toggle_language():
    if current_lang == 'en': set_language('ja')
    else: set_language('en')

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

def preprocess_for_ocr(pil_image, scale_factor=4):
    """Prepares a small PIL image for better OCR results."""
    try:
        cv_img = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
        width = int(cv_img.shape[1] * scale_factor)
        height = int(cv_img.shape[0] * scale_factor)
        if width == 0 or height == 0: return None
        upscaled = cv2.resize(cv_img, (width, height), interpolation=cv2.INTER_LANCZOS4)
        gray = cv2.cvtColor(upscaled, cv2.COLOR_BGR2GRAY)
        bw_img = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                     cv2.THRESH_BINARY, 11, 2)
        bw_img_inverted = cv2.bitwise_not(bw_img)
        return bw_img_inverted
    except Exception as e:
        print(f"OCR Preprocessing error: {e}")
        return cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2GRAY)

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
    
    for key, value in data_results.items():
        if value in ["N/A", "Error"]:
            is_complete = False
            break
            
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
    """Calls validation and GSheet sender."""
    try:
        image = ImageGrab.grab()
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
                status_text, status_color = validate_data(data_results)
                if match_name != "None" and status_text == translations['validation_pass'][current_lang]:
                    send_data_to_google_sheet(match_name, data_results)
                final_results.append((crop_pil, match_name, (x, y), data_results, (status_text, status_color)))
        else:
            crop_pil = image
            match_name = find_best_tabname_match(crop_pil)
            data_results = {}
            if match_name != "None":
                data_results = extract_data_from_rois(crop_pil, match_name, (0, 0))
            status_text, status_color = validate_data(data_results)
            if match_name != "None" and status_text == translations['validation_pass'][current_lang]:
                send_data_to_google_sheet(match_name, data_results)
            final_results.append((crop_pil, match_name, (0, 0), data_results, (status_text, status_color)))

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
        
    for roi_key, [global_x, global_y, global_w, global_h] in rois_to_draw.items():
        try:
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
                processed_img = preprocess_for_ocr(roi_crop_pil)
                if processed_img is None:
                    data_results[roi_key] = "N/A"
                    continue
                ocr_results = ocr_reader.readtext(processed_img, allowlist=OCR_ALLOWLIST, detail=0)
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
    global image_placeholder_label, auto_cap_photos
    auto_cap_photos.clear()
    for widget in crop_display_frame.winfo_children():
        widget.destroy()
    image_placeholder_label = tk.Label(crop_display_frame, font=(font_family, 12, 'italic'), background='#ffffff', foreground='#888888', text=translations['image_placeholder'][current_lang])
    image_placeholder_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)


def update_gui_with_sift_results(sift_results):
    """(REPLACED v0.13) New side-by-side layout."""
    global auto_cap_photos, image_placeholder_label
    
    auto_cap_photos.clear()
    for widget in crop_display_frame.winfo_children():
        widget.destroy()
    image_placeholder_label = None 
    if not sift_results:
        clear_image_display()
        return

    num_images = len(sift_results)
    
    for (pil_image, match_name, (crop_offset_x, crop_offset_y), data_results, (status_text, status_color)) in sift_results:
        
        result_frame = tk.Frame(crop_display_frame, background="#f0f0f0", relief=tk.SUNKEN, borderwidth=1)
        
        image_frame = ttk.Frame(result_frame)
        image_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 10), pady=5)
        
        data_frame = ttk.Frame(result_frame, width=250)
        data_frame.pack(side=tk.LEFT, fill=tk.Y, pady=5, ipadx=10)
        data_frame.pack_propagate(False)

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

        container_width = (root.winfo_width() // num_images) - 300
        if container_width < 100: container_width = 100
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

        result_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)

    now_time = datetime.datetime.now().strftime('%H:%M:%S')
    update_status('status_captured', now_time)

# --- Region Capture & ROI Logic ---
class RegionSelector:
    def __init__(self, parent):
        self.parent = parent
        self.background_image = ImageGrab.grab()
        self.selector_window = tk.Toplevel(parent)
        self.selector_window.attributes('-fullscreen', True)
        self.selector_window.attributes('-alpha', 0.3)
        self.selector_window.wait_visibility(self.selector_window)
        self.canvas = tk.Canvas(self.selector_window, cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.tk_background_image = ImageTk.PhotoImage(self.background_image)
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
        self.box = (int(min(self.start_x, end_x)), int(min(self.start_y, end_y)), int(max(self.start_x, end_x)), int(max(self.start_y, end_y)))
        self.selector_window.destroy()

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
            now = datetime.datetime.now()
            timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"capture_{timestamp}.png"
            save_path = os.path.join(TABNAME_DIR, filename)
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
    set_filename = simpledialog.askstring(
        translations['roi_save_as_title'][current_lang],
        translations['roi_save_as_text'][current_lang],
        parent=root
    )
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
        update_status('status_error', str(e))
        return
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
        new_filename = simpledialog.askstring(
            translations['rename_prompt_title'][current_lang],
            translations['rename_prompt_text'][current_lang],
            initialvalue=old_filename, parent=root
        )
        if new_filename and new_filename != old_filename:
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
        new_filename = simpledialog.askstring(
            translations['rename_prompt_title'][current_lang],
            translations['rename_prompt_text'][current_lang],
            initialvalue=old_filename, parent=root
        )
        if new_filename and new_filename != old_filename:
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
    foldername = simpledialog.askstring(
        translations['create_folder_button'][current_lang],
        translations['rename_prompt_text'][current_lang],
        parent=root
    )
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
        new_foldername = simpledialog.askstring(
            translations['rename_folder_button'][current_lang],
            translations['rename_prompt_text'][current_lang],
            initialvalue=old_foldername, parent=root
        )
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
            imagename = simpledialog.askstring(
                translations['add_image_prompt_title'][current_lang],
                translations['add_image_prompt_text'][current_lang],
                parent=root
            )
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
        new_filename = simpledialog.askstring(
            translations['rename_image_button'][current_lang],
            translations['rename_prompt_text'][current_lang],
            initialvalue=old_filename, parent=root
        )
        if new_filename and new_filename != old_filename:
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
root.geometry("850x650")
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
interval_entry = ttk.Entry(settings_frame, width=5, font=(font_family, 10))
interval_entry.pack(side=tk.LEFT, padx=5)
interval_entry.insert(0, "5") 
start_button = ttk.Button(settings_frame, command=start_capture)
start_button.pack(side=tk.LEFT, padx=5)
stop_button = ttk.Button(settings_frame, command=stop_capture, state=tk.DISABLED)
stop_button.pack(side=tk.LEFT, padx=5)
capture_region_button = ttk.Button(settings_frame, command=start_region_capture)
capture_region_button.pack(side=tk.LEFT, padx=15)
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
crop_display_frame = tk.Frame(capture_tab, background='#ffffff', relief=tk.SUNKEN, borderwidth=1)
crop_display_frame.pack(fill=tk.BOTH, expand=True, pady=5)

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
gallery_image_list.pack(fill=tk.BOTH, expand=True)
gallery_image_list.heading("#0", text="")
gallery_image_list.bind("<<TreeviewSelect>>", on_gallery_item_select)
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
roi_set_list.pack(fill=tk.BOTH, expand=True)
roi_set_list.heading("#0", text="ROI Set Files")
roi_set_list.bind("<<TreeviewSelect>>", on_roi_set_select)
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
status_folder_list.pack(fill=tk.BOTH, expand=True, pady=5)
status_folder_list.heading("#0", text="Tabname Folders")
status_folder_list.bind("<<TreeviewSelect>>", on_status_folder_select)
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
status_image_list.pack(fill=tk.BOTH, expand=True, pady=5)
status_image_list.heading("#0", text="Status Images")
status_image_list.bind("<<TreeviewSelect>>", on_status_image_select)
status_preview_label = tk.Label(status_image_frame, background="#ffffff", relief=tk.SUNKEN, borderwidth=1)
status_preview_label.pack(fill=tk.BOTH, expand=True, pady=5)

# ---- 7. สร้าง Tab 5: Settings ----
settings_tab = ttk.Frame(notebook, padding=10)
notebook.add(settings_tab, text="Settings")

g_sheet_frame = ttk.Frame(settings_tab)
g_sheet_frame.pack(fill=tk.X, pady=10)

g_sheet_url_label = ttk.Label(g_sheet_frame, text="Google Sheet Web App URL:", anchor=tk.W)
g_sheet_url_label.pack(fill=tk.X)

# (MODIFIED) Use the new EntryWithRightClickMenu
g_sheet_url_entry = EntryWithRightClickMenu(g_sheet_frame, width=80) 
g_sheet_url_entry.pack(fill=tk.X, pady=(5, 10))

g_sheet_save_button = ttk.Button(g_sheet_frame, command=save_config)
g_sheet_save_button.pack(anchor=tk.W)


# ---- 8. สร้างแถบสถานะ (ล่างสุด) ----
status_label = ttk.Label(root, relief=tk.SUNKEN, anchor=tk.W, padding=5, font=(font_family, 9))
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# ---- 9. ตั้งค่าการปิดหน้าต่าง และเริ่มแอป ----
root.protocol("WM_DELETE_WINDOW", on_closing) 
set_language(current_lang)
clear_image_display() 
refresh_gallery_list() # This loads ALL SIFT caches
refresh_roi_file_list()
refresh_status_folders()
on_gallery_item_select(None)
on_roi_set_select(None)
load_config() # Load the saved URL
root.mainloop()