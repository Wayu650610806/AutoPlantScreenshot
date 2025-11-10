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

# --- SIFT Global Initialization ---
try:
    sift = cv2.SIFT_create()
    sift_template_cache = {}
    SIFT_MATCH_THRESHOLD = 70
except Exception as e:
    messagebox.showerror("OpenCV Error", f"ไม่สามารถเริ่ม SIFT ได้ (อาจต้องติดตั้ง opencv-contrib-python)\n{e}")
    sys.exit()

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
os.makedirs(TABNAME_DIR, exist_ok=True)
ROI_DIR = os.path.join(BASE_PATH, "rois")
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
    'app_title': {'en': 'Capture Tool v0.10 (File Match)', 'ja': 'キャプチャーツール v0.10 (ファイル一致)'}, # (MODIFIED)
    'tab_capture': {'en': 'Auto-Capture', 'ja': '自動キャプチャ'},
    'tab_gallery': {'en': 'Tabname', 'ja': 'タブ名'}, 
    'tab_roi_sets': {'en': 'ROI Sets', 'ja': 'ROIセット'},
    'interval_label': {'en': 'Interval (sec):', 'ja': '間隔 (秒):'},
    'start_button': {'en': 'Start Auto', 'ja': '自動開始'},
    'stop_button': {'en': 'Stop Auto', 'ja': '自動停止'},
    'progress_label': {'en': 'Next capture in:', 'ja': '次のキャプチャ:'},
    'image_placeholder': {'en': 'Captured crops will appear here\n(Press "Start" to begin)', 'ja': 'キャプチャした画像はここに表示されます\n(「開始」を押してください)'},
    'split_method_label': {'en': 'Split Method:', 'ja': '分割方法:'},
    'capture_region_button': {'en': 'Capture Tabname', 'ja': 'タブ名キャプチャ'},
    # 'active_roi_set_label': DELETED
    'gallery_preview_header': {'en': 'Preview', 'ja': 'プレビュー'},
    'gallery_refresh_button': {'en': 'Refresh', 'ja': '更新'},
    'gallery_rename_button': {'en': 'Rename', 'ja': '名前変更'},
    'gallery_delete_button': {'en': 'Delete', 'ja': '削除'},
    'gallery_placeholder': {'en': 'Select an image to preview', 'ja': '画像を選択してください'},
    'create_roi_set_button': {'en': 'Create New ROI Set', 'ja': 'ROIセット新規作成'},
    'roi_set_list_header': {'en': 'ROI Set Files', 'ja': 'ROIセットファイル'},
    'roi_save_as_title': {'en': 'Save ROI Set As', 'ja': 'ROIセットを名前を付けて保存'},
    'roi_save_as_text': {'en': 'Enter a filename for this new ROI set (e.g., "machine_A.json"):', 'ja': 'このROIセットのファイル名を入力してください (例: "machine_A.json"):'},
    'roi_add_another_title': {'en': 'Add ROI', 'ja': 'ROI追加'},
    'roi_add_another_text': {'en': 'ROI "{content}" added to set. Add another one?', 'ja': 'ROI「{content}」をセットに追加しました。続けて追加しますか？'},
    'roi_name_prompt_title': {'en': 'Enter ROI Name', 'ja': 'ROI名入力'},
    'roi_name_prompt_text': {'en': 'Select or type a name for this ROI:', 'ja': 'このROIの名前を選択または入力してください:'},
    'rename_prompt_title': {'en': 'Rename File', 'ja': '名前の変更'},
    'rename_prompt_text': {'en': 'Enter new filename:', 'ja': '新しいファイル名を入力:'},
    'delete_confirm_title': {'en': 'Confirm Delete', 'ja': '削除の確認'},
    'delete_confirm_text': {'en': 'Are you sure you want to delete this file?\n{content}', 'ja': 'このファイルを削除してもよろしいですか？\n{content}'},
    'status_idle': {'en': 'Status: Idle', 'ja': 'ステータス: 待機中'},
    'status_running': {'en': 'Auto-Capture running... (every {content} sec)', 'ja': '自動キャプチャ実行中... ({content} 秒ごと)'},
    'status_stopped': {'en': 'Status: Stopped', 'ja': 'ステータス: 停止'},
    'status_captured': {'en': 'Last captured: {content}', 'ja': '最終キャプチャ: {content}'},
    'status_error': {'en': 'Error: {content}', 'ja': 'エラー: {content}'},
    'status_saved': {'en': 'Image saved to gallery: {content}', 'ja': 'ギャラリーに画像を保存しました: {content}'},
    'status_sift_loading': {'en': 'Loading SIFT templates...', 'ja': 'SIFTテンプレートを読込中...'},
    'status_sift_done': {'en': 'SIFT templates loaded ({content} images).', 'ja': 'SIFTテンプレートを読込完了 ({content} 画像).'},
    'status_roi_saved': {'en': 'ROI Set "{content}" saved.', 'ja': 'ROIセット「{content}」を保存しました。'},
    # 'status_roi_loaded': DELETED
    'status_roi_error': {'en': 'Failed to load/save ROI data.', 'ja': 'ROIデータの読み込み/保存に失敗しました。'},
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
roi_rename_button = None
roi_delete_button = None
roi_refresh_button = None

# --- Custom Dialog for ROI Naming ---
class AskROINameDialog(simpledialog.Dialog):
    def __init__(self, parent, title, text, predefined_names):
        self.text = text
        self.predefined_names = predefined_names
        super().__init__(parent, title)
    def body(self, master):
        ttk.Label(master, text=self.text).pack(pady=5)
        self.combo = ttk.Combobox(master, values=self.predefined_names, width=50)
        self.combo.pack(padx=10, pady=5)
        return self.combo
    def apply(self):
        self.result = self.combo.get()

# --- (DELETED) load_active_roi_file ---

def set_language(lang_code):
    global current_lang, image_placeholder_label
    current_lang = lang_code
    
    root.title(translations['app_title'][current_lang])
    notebook.tab(capture_tab, text=translations['tab_capture'][current_lang])
    notebook.tab(gallery_tab, text=translations['tab_gallery'][current_lang])
    notebook.tab(roi_tab, text=translations['tab_roi_sets'][current_lang])
    
    # Capture Tab
    interval_label.config(text=translations['interval_label'][current_lang])
    start_button.config(text=translations['start_button'][current_lang])
    stop_button.config(text=translations['stop_button'][current_lang])
    lang_button.config(text=translations['lang_button'][current_lang])
    progress_label.config(text=translations['progress_label'][current_lang])
    split_method_label.config(text=translations['split_method_label'][current_lang])
    capture_region_button.config(text=translations['capture_region_button'][current_lang])
    # (DELETED) active_roi_set_label
    
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
    roi_rename_button.config(text=translations['gallery_rename_button'][current_lang])
    roi_delete_button.config(text=translations['gallery_delete_button'][current_lang])
    roi_refresh_button.config(text=translations['gallery_refresh_button'][current_lang])
        
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

# --- SIFT Matching Logic ---
def pil_to_cv2_gray(pil_image):
    return cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2GRAY)

def load_sift_templates():
    global sift_template_cache
    sift_template_cache.clear()
    update_status('status_sift_loading')
    root.update_idletasks()
    count = 0
    try:
        for filename in os.listdir(TABNAME_DIR):
            if filename.endswith('.png'):
                filepath = os.path.join(TABNAME_DIR, filename)
                img_bytes = np.fromfile(filepath, dtype=np.uint8)
                img = cv2.imdecode(img_bytes, cv2.IMREAD_GRAYSCALE)
                if img is None: continue
                kp, des = sift.detectAndCompute(img, None)
                if des is not None and len(kp) > 0:
                    sift_template_cache[filename] = (kp, des)
                    count += 1
        update_status('status_sift_done', count)
    except Exception as e:
        update_status('status_error', f"SIFT template load failed: {e}")

def find_best_match(cropped_pil_image):
    if not sift_template_cache: return "None"
    try:
        w, h = cropped_pil_image.size
        top_half_box = (0, 0, w, h // 2)
        image_to_check = cropped_pil_image.crop(top_half_box)
        img_crop_gray = pil_to_cv2_gray(image_to_check)
        kp_crop, des_crop = sift.detectAndCompute(img_crop_gray, None)
        if des_crop is None or len(kp_crop) < SIFT_MATCH_THRESHOLD:
            return "None"
        FLANN_INDEX_KDTREE = 1
        index_params = dict(algorithm=FLANN_INDEX_KDTREE, trees=5)
        search_params = dict(checks=50)
        flann = cv2.FlannBasedMatcher(index_params, search_params)
        best_match_name = "None"
        max_good_matches = 0
        for filename, (kp_template, des_template) in sift_template_cache.items():
            if des_template is None: continue
            matches = flann.knnMatch(des_crop, des_template, k=2)
            good_matches = [m for m, n in matches if m.distance < 0.7 * n.distance]
            if len(good_matches) > SIFT_MATCH_THRESHOLD and len(good_matches) > max_good_matches:
                max_good_matches = len(good_matches)
                best_match_name = filename
        return best_match_name
    except Exception as e:
        print(f"SIFT match error: {e}")
        return "None"

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
    # (DELETED) active_roi_combo
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
    # (DELETED) active_roi_combo
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
    """(MODIFIED v0.10) ส่ง (ภาพ, ชื่อ, offset)"""
    try:
        image = ImageGrab.grab()
        selected_index = split_method_combo.current()
        method_key = SPLIT_ORDER[selected_index]
        split_function = SPLIT_OPTIONS[method_key]['func']
        
        sift_results = []
        
        if split_function:
            boxes = split_function(image) # This returns (x, y, w, h)
            for (x, y, w, h) in boxes:
                box_pil = (x, y, x + w, y + h)
                crop_pil = image.crop(box_pil)
                match_name = find_best_match(crop_pil)
                sift_results.append((crop_pil, match_name, (x, y))) 
        else:
            crop_pil = image
            match_name = find_best_match(crop_pil)
            sift_results.append((crop_pil, match_name, (0, 0))) 

        root.after(0, update_gui_with_sift_results, sift_results)
        
    except Exception as e:
        root.after(0, update_status, 'status_error', str(e))

def clear_image_display():
    global image_placeholder_label, auto_cap_photos
    auto_cap_photos.clear()
    for widget in crop_display_frame.winfo_children():
        widget.destroy()
    image_placeholder_label = tk.Label(crop_display_frame, font=(font_family, 12, 'italic'), background='#ffffff', foreground='#888888', text=translations['image_placeholder'][current_lang])
    image_placeholder_label.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

def update_gui_with_sift_results(sift_results):
    """
    (REPLACED v0.10) วาด ROI ทั้งหมดจากไฟล์ JSON ที่ชื่อตรงกับ SIFT Match
    """
    global auto_cap_photos, image_placeholder_label
    
    auto_cap_photos.clear()
    for widget in crop_display_frame.winfo_children():
        widget.destroy()
    image_placeholder_label = None 
    if not sift_results:
        clear_image_display()
        return

    num_images = len(sift_results)
    container_width = crop_display_frame.winfo_width()
    container_height = crop_display_frame.winfo_height()
    if container_width <= 1: container_width = 830
    if container_height <= 1: container_height = 450
    max_img_width = (container_width // num_images) - (num_images * 4) 
    max_img_height = container_height - 40

    for (pil_image, match_name, (crop_offset_x, crop_offset_y)) in sift_results:
        result_frame = tk.Frame(crop_display_frame, background="#ffffff", relief=tk.SUNKEN, borderwidth=1)
        display_image = pil_image.copy()
        
        # --- (NEW LOGIC v0.10) ---
        match_name_no_ext = match_name.replace(".png", "")
        roi_filename = match_name_no_ext + ".json"
        roi_filepath = os.path.join(ROI_DIR, roi_filename)
        
        # 1. ตรวจสอบว่ามีไฟล์ ROI ที่ชื่อตรงกันหรือไม่
        if os.path.exists(roi_filepath):
            try:
                # 2. โหลดไฟล์ ROI นั้น
                with open(roi_filepath, 'r', encoding='utf-8') as f:
                    rois_to_draw = json.load(f)
                
                # 3. แปลงเป็น CV2 (ครั้งเดียว)
                cv_image = cv2.cvtColor(np.array(display_image), cv2.COLOR_RGB2BGR)

                # 4. วนลูปวาด *ทุกกล่อง* ในไฟล์นั้น
                for roi_key, [global_x, global_y, global_w, global_h] in rois_to_draw.items():
                    # 5. แก้ไขพิกัด (บั๊กที่เราคุยกัน)
                    local_x = global_x - crop_offset_x
                    local_y = global_y - crop_offset_y
                    
                    # 6. วาดกล่อง
                    cv2.rectangle(cv_image, (local_x, local_y), (local_x + global_w, local_y + global_h), (0, 255, 0), 2)
                
                # 7. แปลงกลับเป็น PIL
                display_image = Image.fromarray(cv2.cvtColor(cv_image, cv2.COLOR_BGR2RGB))
            
            except Exception as e:
                print(f"Error loading/drawing ROI file {roi_filename}: {e}")
                # (ถ้าพลาด ก็ใช้ภาพเดิมต่อไป)
        # --- (END NEW LOGIC) ---

        display_image.thumbnail((max_img_width, max_img_height), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(display_image)
        auto_cap_photos.append(photo)
        image_label = tk.Label(result_frame, image=photo, background="#ffffff")
        image_label.image = photo 
        image_label.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=2, pady=2)

        display_name = match_name.replace(".png", "")
        name_color = "green" if match_name != "None" else "red"
        name_label = tk.Label(result_frame, text=display_name, font=(font_family, 10, 'bold'), background="#ffffff", foreground=name_color)
        name_label.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 5))

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

def start_roi_set_creation():
    """Called by 'Create New ROI Set' button."""
    new_rois = {}
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
            new_rois[roi_name] = [x, y, w, h]
            if not messagebox.askyesno(
                translations['roi_add_another_title'][current_lang],
                translations['roi_add_another_text'][current_lang].format(content=roi_name)
            ):
                break
        else:
            if not messagebox.askyesno("Cancel?", "No name entered. Stop creating this ROI set?"):
                continue
            else:
                break
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
                json.dump(new_rois, f, indent=4, ensure_ascii=False) # Fix for Japanese
            update_status('status_roi_saved', set_filename)
            refresh_roi_file_list()
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
        load_sift_templates()
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
    """(MODIFIED) Refreshes ROI Set list ONLY."""
    try:
        for item in roi_set_list.get_children():
            roi_set_list.delete(item)
        files = os.listdir(ROI_DIR)
        json_files = sorted([f for f in files if f.endswith('.json')], reverse=True)
        for filename in json_files:
            roi_set_list.insert("", tk.END, text=filename, iid=filename)
        # (DELETED) Dropdown update logic
    except Exception as e:
        update_status('status_error', str(e))

# (DELETED) on_active_roi_select function

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

# ---- 2. สร้าง Notebook (Tabbed Interface) ----
notebook = ttk.Notebook(root, padding=10)
notebook.pack(fill=tk.BOTH, expand=True)

# ---- 3. สร้าง Tab 1: Auto-Capture ----
capture_tab = ttk.Frame(notebook)
notebook.add(capture_tab, text="Capture")

# -- Settings Frame --
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

# -- Split Frame --
split_frame = ttk.Frame(capture_tab, padding=(0, 8, 0, 0))
split_frame.pack(fill=tk.X)
split_method_label = ttk.Label(split_frame)
split_method_label.pack(side=tk.LEFT, padx=(0, 5))
split_method_combo = ttk.Combobox(split_frame, state="readonly", font=(font_family, 10))
split_method_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

# -- (DELETED) Active ROI Set Frame was here --

# -- Progress Bar Frame --
progress_frame = ttk.Frame(capture_tab, padding=(0, 10, 0, 5))
progress_frame.pack(fill=tk.X)
progress_label = ttk.Label(progress_frame)
progress_label.pack(side=tk.LEFT, padx=(0, 5))
progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode='determinate')
progress_bar.pack(fill=tk.X, expand=True)

# -- Crop Display Frame --
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
roi_refresh_button = ttk.Button(roi_action_frame, command=refresh_roi_file_list)
roi_refresh_button.pack(side=tk.LEFT, padx=5, pady=5)
roi_paned_window = ttk.PanedWindow(roi_tab, orient=tk.HORIZONTAL)
roi_paned_window.pack(fill=tk.BOTH, expand=True)
roi_list_frame = ttk.Frame(roi_paned_window, padding=5)
roi_paned_window.add(roi_list_frame, weight=1)
roi_set_list = ttk.Treeview(roi_list_frame, selectmode="browse")
roi_set_list.pack(fill=tk.BOTH, expand=True)
roi_set_list.heading("#0", text="ROI Set Files")
roi_edit_frame = ttk.Frame(roi_paned_window, padding=10)
roi_paned_window.add(roi_edit_frame, weight=1)
ttk.Label(roi_edit_frame, text="Selected File Actions:", font=(font_family, 11, 'bold')).pack(pady=10)
roi_rename_button = ttk.Button(roi_edit_frame, command=rename_roi_file)
roi_rename_button.pack(fill=tk.X, padx=5, pady=5)
roi_delete_button = ttk.Button(roi_edit_frame, command=delete_roi_file)
roi_delete_button.pack(fill=tk.X, padx=5, pady=5)

# ---- 6. สร้างแถบสถานะ (ล่างสุด) ----
status_label = ttk.Label(root, relief=tk.SUNKEN, anchor=tk.W, padding=5, font=(font_family, 9))
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# ---- 7. ตั้งค่าการปิดหน้าต่าง และเริ่มแอป ----
root.protocol("WM_DELETE_WINDOW", on_closing) 
set_language(current_lang)
clear_image_display() 
refresh_gallery_list()
refresh_roi_file_list()
on_gallery_item_select(None)
root.mainloop()