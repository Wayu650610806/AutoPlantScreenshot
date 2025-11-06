import time
import schedule
import datetime
from PIL import ImageGrab

def capture_job():
    """
    ฟังก์ชันหลักที่จะทำงานตามตารางเวลา
    """
    try:
        # 1. ดึงภาพหน้าจอทั้งหมด
        image = ImageGrab.grab()
        
        # 2. สร้างชื่อไฟล์ที่ไม่ซ้ำกัน
        # เช่น "capture_2025-11-06_09-30-15.png"
        now = datetime.datetime.now()
        timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"capture_{timestamp}.png"
        
        # 3. บันทึกไฟล์ลงดิสก์
        image.save(filename)
        
        print(f"[{now.strftime('%H:%M:%S')}] บันทึกภาพหน้าจอ: {filename}")
        
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการแคปหน้าจอ: {e}")

# --- ส่วนตั้งค่าการทำงาน ---

# ตั้งตารางเวลา
# เปลี่ยน .seconds เป็น .minutes ได้ถ้าต้องการ
schedule.every(10).seconds.do(capture_job)

print("--- เริ่มโปรแกรมแคปหน้าจออัตโนมัติ ---")
print("จะทำการบันทึกภาพหน้าจอทุกๆ 10 วินาที")
print("กด Ctrl+C เพื่อหยุดการทำงาน")

# สั่งให้โปรแกรมทำงานวนลูปเพื่อตรวจสอบตารางเวลา
try:
    while True:
        schedule.run_pending()
        time.sleep(1)
except KeyboardInterrupt:
    print("\nกำลังหยุดการทำงาน...")
    print("--- โปรแกรมปิดตัวลง ---")