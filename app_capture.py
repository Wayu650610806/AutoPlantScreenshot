import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from PIL import ImageGrab, ImageTk, Image
import threading
import datetime
import sys

# --- DPI Awareness Fix (แก้ปัญหาตัวอักษรแตกบน Windows) ---
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1) 
except ImportError:
    pass 
except Exception as e:
    print(f"DPI Awareness Error: {e}")

# --- (NEW) Split Functions (ปรับจากโค้ดของคุณ) ---
# ผมเปลี่ยนชื่อฟังก์ชันไม่ให้ซ้ำกัน และเปลี่ยน .shape เป็น .size
# ทั้งหมดนี้รับ PIL.Image และคืนค่าเป็น list of (x, y, w, h)

def split_pattern_1(pil_image):
    """แบ่งภาพ 34%, 34%, 16%, 16%"""
    w, h = pil_image.size
    split_x_1 = (w * 34) // 100
    split_x_2 = (w * 68) // 100
    split_x_3 = (w * 84) // 100
    boxes = []
    
    w1 = split_x_1
    boxes.append((0, 0, w1, h))

    x2 = split_x_1
    w2 = split_x_2 - split_x_1 
    boxes.append((x2, 0, w2, h))

    x3 = split_x_2
    w3 = split_x_3 - split_x_2
    boxes.append((x3, 0, w3, h))
    
    x4 = split_x_3
    w4 = w - split_x_3
    boxes.append((x4, 0, w4, h))
    return boxes

def split_pattern_2(pil_image):
    """แบ่งภาพ 4 ส่วนเท่าๆ กัน (25% x 4)"""
    w, h = pil_image.size
    split_x_1 = w // 4
    split_x_2 = w // 2
    split_x_3 = (w * 3) // 4
    boxes = []
    
    w1 = split_x_1
    boxes.append((0, 0, w1, h))

    x2 = split_x_1
    w2 = split_x_2 - split_x_1 
    boxes.append((x2, 0, w2, h))

    x3 = split_x_2
    w3 = split_x_3 - split_x_2
    boxes.append((x3, 0, w3, h))
    
    x4 = split_x_3
    w4 = w - split_x_3
    boxes.append((x4, 0, w4, h))
    return boxes

def split_pattern_3(pil_image):
    """แบ่งภาพ 3 ส่วน: 25%, 25%, 50%"""
    w, h = pil_image.size
    split_x_1 = w // 2
    split_x_2 = split_x_1 // 2
    boxes = []
    
    w1 = split_x_2
    boxes.append((0, 0, w1, h))

    x2 = split_x_2
    w2 = split_x_1 - split_x_2 
    boxes.append((x2, 0, w2, h))

    x3 = split_x_1
    w3 = w - split_x_1 
    boxes.append((x3, 0, w3, h))
    return boxes

# --- (NEW) Global list of split method keys ---
# เราจะใช้ key เหล่านี้ในการอ้างอิงและแปลภาษา
SPLIT_OPTIONS = {
    "NONE": {
        'en': 'None (Full Image)',
        'ja': 'なし (フルイメージ)',
        'func': None
    },
    "P1_34_34_16_16": {
        'en': 'Pattern 1 (34/34/16/16)',
        'ja': 'パターン1 (34/34/16/16)',
        'func': split_pattern_1
    },
    "P2_25x4": {
        'en': 'Pattern 2 (25% x 4)',
        'ja': 'パターン2 (25% x 4)',
        'func': split_pattern_2
    },
    "P3_25_25_50": {
        'en': 'Pattern 3 (25/25/50)',
        'ja': 'パターン3 (25/25/50)',
        'func': split_pattern_3
    }
}
# แปลงเป็น list ของ keys เพื่อรักษาลำดับ
SPLIT_ORDER = ["NONE", "P1_34_34_16_16", "P2_25x4", "P3_25_25_50"]


# --- Language Data (คลังคำศัพท์) ---
translations = {
    'app_title': {'en': 'Screen Capture v0.5 (Splitter)', 'ja': 'スクリーンキャプチャ v0.5 (分割)'},
    'interval_label': {'en': 'Interval (sec):', 'ja': '間隔 (秒):'},
    'start_button': {'en': 'Start', 'ja': '開始'},
    'stop_button': {'en': 'Stop', 'ja': '停止'},
    'lang_button': {'en': '日本語', 'ja': 'English'},
    'progress_label': {'en': 'Next capture in:', 'ja': '次のキャプチャ:'},
    'image_placeholder': {'en': 'Captured crops will appear here\n(Press "Start" to begin)', 'ja': 'キャプチャした画像はここに表示されます\n(「開始」を押してください)'},
    'status_idle': {'en': 'Status: Idle', 'ja': 'ステータス: 待機中'},
    'status_running': {'en': 'Starting... (every {content} sec)', 'ja': '開始中... ({content} 秒ごと)'},
    'status_stopped': {'en': 'Status: Stopped', 'ja': 'ステータス: 停止'},
    'status_captured': {'en': 'Last captured: {content}', 'ja': '最終キャプチャ: {content}'},
    'status_error': {'en': 'Error: {content}', 'ja': 'エラー: {content}'},
    'error_title': {'en': 'Invalid Input', 'ja': '無効な入力'},
    'error_message': {'en': 'Please enter a valid number (seconds > 1).\n{content}', 'ja': '有効な数字（1秒以上）を入力してください。\n{content}'},
    'confirm_close_title': {'en': 'Confirm Exit', 'ja': '終了確認'},
    'confirm_close_message': {'en': 'The program is running. Are you sure you want to stop and exit?', 'ja': 'プログラムが実行中です。停止して終了しますか？'},
    
    # (NEW) คำแปลสำหรับส่วนตัดแบ่ง
    'split_method_label': {'en': 'Split Method:', 'ja': '分割方法:'},
    # เพิ่มคำแปลสำหรับแต่ละ key ใน SPLIT_OPTIONS
    **{key: value for key, value in SPLIT_OPTIONS.items()} 
}

# --- Global Variables ---
is_running = False      
timer_job_id = None     
current_lang = 'en' 
photo_images = [] # (NEW) List นี้สำคัญมาก! ใช้เก็บ reference ของ PhotoImage กันภาพหาย
image_placeholder_label = None # (NEW) ตัวแปรสำหรับเก็บ Label placeholder

def set_language(lang_code):
    """อัปเดต UI ทั้งหมดเป็นภาษาที่กำหนด"""
    global current_lang, image_placeholder_label
    current_lang = lang_code
    
    root.title(translations['app_title'][current_lang])
    interval_label.config(text=translations['interval_label'][current_lang])
    start_button.config(text=translations['start_button'][current_lang])
    stop_button.config(text=translations['stop_button'][current_lang])
    lang_button.config(text=translations['lang_button'][current_lang])
    progress_label.config(text=translations['progress_label'][current_lang])
    split_method_label.config(text=translations['split_method_label'][current_lang]) # (NEW)

    # (NEW) อัปเดตค่าใน Combobox
    current_index = split_method_combo.current()
    if current_index == -1: current_index = 0
    
    split_display_options = [SPLIT_OPTIONS[key][current_lang] for key in SPLIT_ORDER]
    split_method_combo.config(values=split_display_options)
    split_method_combo.current(current_index)
    
    # อัปเดตข้อความสถานะและ placeholder
    if not is_running:
        status_label.config(text=translations['status_idle'][current_lang])
        if image_placeholder_label:
            image_placeholder_label.config(text=translations['image_placeholder'][current_lang])
    else:
        interval = interval_entry.get()
        update_status('status_running', interval)


def toggle_language():
    """สลับภาษาระหว่าง EN และ JA"""
    if current_lang == 'en':
        set_language('ja')
    else:
        set_language('en')

def start_capture():
    global is_running, timer_job_id
    if is_running: return

    try:
        interval = int(interval_entry.get())
        if interval <= 1:
            raise ValueError("Time must be > 1")
    except ValueError as e:
        messagebox.showerror(
            translations['error_title'][current_lang], 
            translations['error_message'][current_lang].format(content=e)
        )
        return

    is_running = True
    update_status('status_running', interval)
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)
    interval_entry.config(state=tk.DISABLED)
    lang_button.config(state=tk.DISABLED)
    split_method_combo.config(state=tk.DISABLED) # (NEW) ล็อคการเลือก
    
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
    split_method_combo.config(state=tk.NORMAL) # (NEW) ปลดล็อค
    
    progress_bar['value'] = 0
    update_status('status_stopped')
    clear_image_display() # (NEW) ล้างหน้าจอเมื่อหยุด

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
    """(MODIFIED) แคป, ตัดแบ่ง, และส่ง list ของภาพไปอัปเดต GUI"""
    try:
        # 1. แคปภาพ
        image = ImageGrab.grab()
        
        # 2. (NEW) หาวิธีตัดแบ่งที่เลือก
        # เราต้องดึงค่า index จาก GUI ใน Main thread
        selected_index = split_method_combo.current()
        method_key = SPLIT_ORDER[selected_index]
        split_function = SPLIT_OPTIONS[method_key]['func']
        
        cropped_images = []
        
        if split_function:
            # 3. (NEW) เรียกใช้ฟังก์ชันตัดแบ่ง
            boxes = split_function(image) # ได้ list of (x, y, w, h)
            
            # 4. (NEW) ตัดภาพ (Crop) จริงๆ
            for (x, y, w, h) in boxes:
                box_pil = (x, y, x + w, y + h) # แปลงเป็น (left, upper, right, lower)
                cropped_images.append(image.crop(box_pil))
        else:
            # ถ้าเลือก "None" ก็เพิ่มภาพเต็ม
            cropped_images.append(image)
            
        # 5. ส่ง list ของภาพที่ตัดแล้วไปอัปเดต GUI
        root.after(0, update_gui_with_cropped_images, cropped_images)
        
    except Exception as e:
        root.after(0, update_status, 'status_error', str(e))

def clear_image_display():
    """(NEW) ล้างภาพทั้งหมดและแสดง placeholder"""
    global image_placeholder_label, photo_images
    
    # ล้าง reference เก่า
    photo_images.clear()
    
    # ทำลาย widget รูปภาพเก่าทั้งหมด
    for widget in crop_display_frame.winfo_children():
        widget.destroy()
        
    # สร้าง placeholder ขึ้นมาใหม่
    image_placeholder_label = tk.Label(crop_display_frame, 
                       font=(font_family, 12, 'italic'), 
                       background='#ffffff', 
                       foreground='#888888',
                       text=translations['image_placeholder'][current_lang])
    image_placeholder_label.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)


def update_gui_with_cropped_images(cropped_images):
    """
    (REPLACED) อัปเดตหน้าจอด้วย *list* ของภาพที่ถูกตัด
    """
    global photo_images, image_placeholder_label
    
    # 1. ล้างหน้าจอเก่า (ทั้ง placeholder หรือรูปเก่า)
    photo_images.clear()
    for widget in crop_display_frame.winfo_children():
        widget.destroy()
    image_placeholder_label = None # ล้างตัวแปร
    
    if not cropped_images:
        clear_image_display() # ถ้ามีอะไรผิดพลาด ให้กลับไปหน้า placeholder
        return

    # 2. คำนวณขนาดที่เหมาะสม
    num_images = len(cropped_images)
    container_width = crop_display_frame.winfo_width()
    container_height = crop_display_frame.winfo_height()
    
    if container_width <= 1: container_width = 580 
    if container_height <= 1: container_height = 350
    
    # ให้ความกว้างเป็นตัวกำหนดหลัก
    max_img_width = (container_width // num_images) - (num_images * 4) 
    max_img_height = container_height - 10

    # 3. วนลูปสร้าง Label รูปภาพใหม่
    for pil_image in cropped_images:
        display_image = pil_image.copy()
        
        # ย่อภาพโดยคงสัดส่วน (ใช้ thumbnail จะง่าย)
        display_image.thumbnail((max_img_width, max_img_height), Image.Resampling.LANCZOS)
        
        photo = ImageTk.PhotoImage(display_image)
        photo_images.append(photo) # *สำคัญ* เก็บ reference ใน list global
        
        # สร้าง Label ใหม่สำหรับภาพนี้
        label = tk.Label(crop_display_frame, image=photo, background="#ffffff", relief=tk.SUNKEN, borderwidth=1)
        label.image = photo # เก็บ reference ในตัว Label ด้วย
        
        # วางเรียงไปทางซ้าย
        label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)

    # 4. อัปเดตสถานะ
    now_time = datetime.datetime.now().strftime('%H:%M:%S')
    update_status('status_captured', now_time)


def update_status(text_key, dynamic_content=""):
    """อัปเดตสถานะโดยใช้ key จาก translations dict"""
    base_text = translations.get(text_key, {}).get(current_lang, text_key)
    final_text = base_text.format(content=dynamic_content)
    status_label.config(text=final_text)

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
root.geometry("600x500") 
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
style.configure('TCombobox', font=(font_family, 10)) # (NEW)

# ---- 2. สร้าง Frame หลัก ----
main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill=tk.BOTH, expand=True)

# ---- 3. สร้าง Frame สำหรับการตั้งค่า (แถวบน) ----
settings_frame = ttk.Frame(main_frame)
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

lang_button = ttk.Button(settings_frame, command=toggle_language)
lang_button.pack(side=tk.RIGHT, padx=5) 

# ---- (NEW) 3.1. สร้าง Frame สำหรับเลือกวิธีตัด (แถว 1.5) ----
split_frame = ttk.Frame(main_frame, padding=(0, 8, 0, 0))
split_frame.pack(fill=tk.X)

split_method_label = ttk.Label(split_frame)
split_method_label.pack(side=tk.LEFT, padx=(0, 5))

# สร้าง Combobox (Dropdown)
split_method_combo = ttk.Combobox(split_frame, state="readonly", font=(font_family, 10))
split_method_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

# ---- 4. สร้าง Frame สำหรับ Countdown Bar (แถวสอง) ----
progress_frame = ttk.Frame(main_frame, padding=(0, 10, 0, 5))
progress_frame.pack(fill=tk.X)

progress_label = ttk.Label(progress_frame)
progress_label.pack(side=tk.LEFT, padx=(0, 5))
progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode='determinate')
progress_bar.pack(fill=tk.X, expand=True)

# ---- 5. สร้าง Frame สำหรับแสดงภาพ (ตรงกลาง) ----
# (MODIFIED) นี่คือ Container ที่จะใช้บรรจุภาพที่ถูกตัด
crop_display_frame = tk.Frame(main_frame, background='#ffffff', relief=tk.SUNKEN, borderwidth=1)
crop_display_frame.pack(fill=tk.BOTH, expand=True, pady=5)
# (เราจะสร้าง placeholder label ตอนที่เรียก set_language ครั้งแรก)


# ---- 6. สร้างแถบสถานะ (ล่างสุด) ----
status_label = ttk.Label(root, relief=tk.SUNKEN, anchor=tk.W, padding=5, font=(font_family, 9))
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# ---- 7. ตั้งค่าการปิดหน้าต่าง และเริ่มแอป ----
root.protocol("WM_DELETE_WINDOW", on_closing) 
set_language(current_lang) # เรียกใช้เพื่อตั้งค่าภาษาเริ่มต้น
clear_image_display() # (NEW) สร้าง placeholder เริ่มต้น
root.mainloop()