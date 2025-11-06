import tkinter as tk
from tkinter import ttk  # ใช้สำหรับ widget ที่สวยขึ้น
from tkinter import messagebox
from PIL import ImageGrab, ImageTk, Image
import threading
import time
import datetime

# --- ตัวแปร Global สำหรับควบคุมการทำงาน ---
is_running = False      # ตัวแปรธง (Flag) เพื่อบอกว่ากำลังทำงานหรือไม่
capture_thread = None   # ตัวแปรสำหรับเก็บ Thread ที่ทำงาน

def start_capture():
    """ฟังก์ชันเมื่อกดปุ่ม 'เริ่ม'"""
    global is_running, capture_thread
    
    if is_running:
        messagebox.showwarning("ทำงานอยู่", "โปรแกรมกำลังทำงานอยู่แล้ว")
        return

    try:
        # 1. อ่านค่าเวลาจากช่อง
        interval = int(interval_entry.get())
        if interval <= 0:
            raise ValueError("เวลต้องมากกว่า 0")
    except ValueError as e:
        messagebox.showerror("ข้อมูลผิดพลาด", f"กรุณาใส่ตัวเลข (วินาที) ให้ถูกต้อง\n{e}")
        return

    # 2. ตั้งค่าสถานะและปุ่ม
    is_running = True
    update_status(f"กำลังเริ่ม... (ทุก {interval} วินาที)")
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)
    interval_entry.config(state=tk.DISABLED)
    
    # 3. เริ่มการทำงานใน Thread ใหม่ (เพื่อไม่ให้หน้าแอปค้าง)
    capture_thread = threading.Thread(target=capture_loop, args=(interval,), daemon=True)
    capture_thread.start()

def stop_capture():
    """ฟังก์ชันเมื่อกดปุ่ม 'หยุด'"""
    global is_running
    is_running = False  # ส่งสัญญาณให้ Thread หยุด
    
    start_button.config(state=tk.NORMAL)
    stop_button.config(state=tk.DISABLED)
    interval_entry.config(state=tk.NORMAL)
    update_status("หยุดทำงานแล้ว")

def capture_loop(interval):
    """
    ฟังก์ชันนี้จะทำงานใน Background Thread
    วนลูปแคปหน้าจอตามเวลาที่กำหนด
    """
    while is_running:
        try:
            # 1. แคปหน้าจอ
            image = ImageGrab.grab()
            
            # 2. บันทึกไฟล์ (เหมือนเดิม)
            now = datetime.datetime.now()
            timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"capture_{timestamp}.png"
            image.save(filename)
            
            # 3. ย่อภาพเพื่อแสดงผลในแอป (ไม่ให้ใหญ่เกินไป)
            display_image = image.copy()
            display_image.thumbnail((400, 300)) # ย่อขนาด (คงอัตราส่วน)
            
            # 4. ส่งภาพและสถานะกลับไปอัปเดตหน้า GUI
            # ต้องใช้ .after() เพื่อความปลอดภัยในการทำงานข้าม Thread
            root.after(0, update_image_display, display_image)
            root.after(0, update_status, f"บันทึกภาพ: {filename}")
            
        except Exception as e:
            root.after(0, update_status, f"เกิดข้อผิดพลาด: {e}")

        # 5. หน่วงเวลา (วิธีนี้จะทำให้ปุ่ม 'หยุด' ตอบสนองได้เร็ว)
        for _ in range(interval):
            if not is_running: # ตรวจสอบธง 'หยุด' ทุกวินาที
                break
            time.sleep(1)

def update_image_display(pil_image):
    """
    (ทำงานใน Main Thread) อัปเดตรูปภาพที่แสดงในหน้าแอป
    """
    # ต้องเก็บ reference ของ PhotoImage ไว้ ไม่งั้นภาพจะหาย (Bug ของ Tkinter)
    global photo_image 
    
    photo_image = ImageTk.PhotoImage(pil_image)
    image_label.config(image=photo_image)
    image_label.image = photo_image

def update_status(text):
    """(ทำงานใน Main Thread) อัปเดตข้อความสถานะ"""
    status_label.config(text=text)

def on_closing():
    """ฟังก์ชันเมื่อผู้ใช้กดปุ่ม X ปิดหน้าต่าง"""
    if is_running:
        if messagebox.askyesno("ยืนยันการปิด", "โปรแกรมกำลังทำงานอยู่ คุณต้องการหยุดและปิดโปรแกรมใช่ไหม?"):
            stop_capture() # หยุด Thread ก่อน
            root.destroy()   # แล้วค่อยปิดหน้าต่าง
    else:
        root.destroy()

# ---- 1. สร้างหน้าต่างหลัก ----
root = tk.Tk()
root.title("โปรแกรมแคปหน้าจออัตโนมัติ v0.2")
root.geometry("450x450") # ขนาดหน้าต่าง กว้างxสูง

# ---- 2. สร้าง Frame สำหรับการตั้งค่า (แถวบน) ----
settings_frame = ttk.Frame(root, padding=10)
settings_frame.pack(fill=tk.X) # ขยายเต็มความกว้าง

ttk.Label(settings_frame, text="ตั้งเวลา (วินาที):").pack(side=tk.LEFT, padx=5)
interval_entry = ttk.Entry(settings_frame, width=5)
interval_entry.pack(side=tk.LEFT, padx=5)
interval_entry.insert(0, "10") # ค่าเริ่มต้น 10 วินาที

start_button = ttk.Button(settings_frame, text="เริ่ม", command=start_capture)
start_button.pack(side=tk.LEFT, padx=5)

stop_button = ttk.Button(settings_frame, text="หยุด", command=stop_capture, state=tk.DISABLED)
stop_button.pack(side=tk.LEFT, padx=5)

# ---- 3. สร้าง Frame สำหรับแสดงภาพ (ตรงกลาง) ----
image_frame = ttk.Frame(root, padding=10)
image_frame.pack(fill=tk.BOTH, expand=True)

# สร้าง Label ว่างๆ ไว้ก่อนเพื่อรอรับภาพ
image_label = ttk.Label(image_frame, text="ยังไม่มีภาพ (กด 'เริ่ม' เพื่อทำงาน)")
image_label.pack(fill=tk.BOTH, expand=True)

# ---- 4. สร้างแถบสถานะ (ล่างสุด) ----
status_label = ttk.Label(root, text="สถานะ: ว่าง", relief=tk.SUNKEN, anchor=tk.W)
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# ---- 5. ตั้งค่าการปิดหน้าต่าง และเริ่มแอป ----
root.protocol("WM_DELETE_WINDOW", on_closing) # ดักจับการกดปุ่ม X
root.mainloop() # เริ่มการทำงานของแอป