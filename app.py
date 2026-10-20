import os
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, session
import openpyxl
from openpyxl import Workbook
try:
    from filelock import FileLock  # นำเข้า FileLock สำหรับจัดการการเขียนไฟล์พร้อมกัน
except ImportError:
    import threading
    class FileLock:
        _lock = threading.Lock()
        def __init__(self, *args, **kwargs):
            pass
        def __enter__(self):
            self._lock.acquire()
            return self
        def __exit__(self, exc_type, exc_val, exc_tb):
            self._lock.release()
from functools import wraps    # นำเข้า wraps สำหรับสร้าง decorator ล็อกอิน

app = Flask(__name__)

# [เปลี่ยน] ย้าย secret_key และรหัสผ่านไปใช้ Environment Variable หากไม่มีให้ใช้ค่าเริ่มต้น (สำหรับตอนทดสอบรัน)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'default_fallback_secret_key')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'CarBooking.xlsx')
LOCK_FILE = os.path.join(BASE_DIR, 'CarBooking.xlsx.lock') # สร้างไฟล์ lock
excel_lock = FileLock(LOCK_FILE, timeout=10) # ตั้งเวลา timeout 10 วินาทีเพื่อรอถ้าไฟล์กำลังถูกเขียน

def init_excel():
    with excel_lock: # [เพิ่ม] ใช้ FileLock ครอบการทำงานที่ยุ่งกับ Excel
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            
            # Sheet 1: Bookings
            ws_bookings = wb.active
            ws_bookings.title = 'Bookings'
            ws_bookings.append(['Booking ID', 'วันที่', 'เวลาเริ่ม', 'เวลาสิ้นสุด', 'รถ', 'ผู้จอง', 'สถานที่', 'วัตถุประสงค์', 'จำนวนผู้เดินทาง', 'หมายเหตุ', 'สถานะ'])
            
            # เพิ่มข้อมูลตัวอย่าง
            ws_bookings.append(['BK001', '2026-10-05', '08:30', '12:00', 'รถ เกี่ยวข้าว', 'นาย A', 'พื้นที่ตนเอง', 'เก็บเกี่ยว', 15, '-', 'รออนุมัติ'])
            ws_bookings.append(['BK002', '2026-10-05', '13:00', '16:30', 'รถ เกี่ยวอ้อย', 'นาย B', 'พื้นที่ตนเอง', 'เก็บเกี่ยว', 10, '-', 'อนุมัติ'])
            ws_bookings.append(['BK003', '2026-10-06', '09:00', '15:00', 'รถ ขุดมันสำปะหลัง', 'นาย C', 'ระยอง', 'เก็บเกี่ยว', 3, '-', 'ไม่อนุมัติ'])
            
            # Sheet 2: Vehicles
            ws_vehicles = wb.create_sheet(title='Vehicles')
            ws_vehicles.append(['รหัสรถ', 'ชนิดของรถ', 'ทะเบียน', 'ที่นั่ง'])
            ws_vehicles.append(['V001', 'รถ เกี่ยวข้าว', 'นข 1234', 1])
            ws_vehicles.append(['V002', 'รถ เกี่ยวอ้อย', 'นข 5678', 1])
            ws_vehicles.append(['V003', 'รถ ขุดมันสำปะหลัง', 'กข 1111', 1])
            ws_vehicles.append(['V004', 'รถ พ่นยา', 'ญณ 1291', 2])
            
            # Sheet 3: Users
            ws_users = wb.create_sheet(title='Users')
            ws_users.append(['รหัสผู้ใช้', 'ชื่อ-สกุล', 'บทบาท'])
            ws_users.append(['U001', 'นาย A', 'ผู้ใช้'])
            ws_users.append(['U002', 'นาย B', 'ผู้ใช้'])
            ws_users.append(['U003', 'Admin', 'ผู้อนุมัติ'])
            
            # Sheet 4: Logs
            ws_logs = wb.create_sheet(title='Logs')
            ws_logs.append(['วันที่/เวลา', 'การดำเนินการ', 'รายละเอียด', 'ผู้ทำ'])
            ws_logs.append(['2026-10-05 10:00', 'สร้างคำขอจอง', 'BK001', 'นาย A'])
            
            wb.save(EXCEL_FILE)

def log_action(action, detail, user):
    with excel_lock: # [เพิ่ม] File lock เวลาบันทึก Log
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_logs = wb['Logs']
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ws_logs.append([now_str, action, detail, user])
        wb.save(EXCEL_FILE)

# [เพิ่ม] Decorator สำหรับบังคับล็อกอินส่วนของ Admin
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in'):
            flash('กรุณาล็อกอินก่อนเข้าใช้งานระบบผู้อนุมัติ', 'warning')
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    return decorated_function

# --- 1. หน้าผู้จอง (User Form) ---
# [เปลี่ยน] ให้ route '/' ทำการ Redirect ไปยัง '/request'
@app.route('/')
def index():
    return redirect(url_for('user_request'))

# [เปลี่ยน] ย้ายหน้าหลักเดิมมาที่ /request
@app.route('/request', methods=['GET', 'POST'])
def user_request():
    if request.method == 'POST':
        date_use = request.form.get('date_use')
        start_time = request.form.get('start_time')
        end_time = request.form.get('end_time')
        vehicle = request.form.get('vehicle')
        location = request.form.get('location')
        purpose = request.form.get('purpose')
        passengers = request.form.get('passengers')
        note = request.form.get('note', '-')
        user_name = request.form.get('user_name')

        with excel_lock: # [เพิ่ม] ใช้ File lock คลุมการอ่านเพื่อตรวจสอบก่อนเขียน
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws_bookings = wb['Bookings']

            # ตรวจสอบว่าช่วงเวลาดังกล่าวถูกจองไปแล้วหรือไม่ (Logic เดิม)
            for row in ws_bookings.iter_rows(min_row=2, values_only=True):
                ex_date = str(row[1])
                ex_start = str(row[2])
                ex_end = str(row[3])
                ex_vehicle = str(row[4])
                ex_status = str(row[10])

                if ex_vehicle == vehicle and ex_date == date_use and ex_status != 'ไม่อนุมัติ':
                    if start_time < ex_end and end_time > ex_start:
                        flash(f'ไม่สามารถจองได้: รถ "{vehicle}" ในวันที่ {date_use} ช่วงเวลา {ex_start} - {ex_end} มีการจองแล้ว!', 'danger')
                        return redirect(url_for('user_request')) # [เปลี่ยน] url_for เป็น user_request

            max_row = ws_bookings.max_row
            booking_id = f"BK{max_row:03d}"

            ws_bookings.append([
                booking_id, date_use, start_time, end_time, 
                vehicle, user_name, location, purpose, passengers, note, 'รออนุมัติ'
            ])
            wb.save(EXCEL_FILE)
        
        log_action('สร้างคำขอจอง', booking_id, user_name)
        flash('ส่งคำขอจองรถเรียบร้อยแล้ว!', 'success')
        return redirect(url_for('user_request')) # [เปลี่ยน] url_for เป็น user_request

    # === ส่วนแสดงผลหน้าเว็บ (GET) ===
    wb = openpyxl.load_workbook(EXCEL_FILE)
    
    ws_vehicles = wb['Vehicles']
    vehicles = [row[1] for row in ws_vehicles.iter_rows(min_row=2, values_only=True)]

    ws_bookings = wb['Bookings']
    busy_list = []
    for row in ws_bookings.iter_rows(min_row=2, values_only=True):
        status = row[10]
        if status in ['อนุมัติ', 'รออนุมัติ']:
            busy_list.append({
                'vehicle': row[4],
                'date': row[1],
                'time': f"{row[2]} - {row[3]}",
                'status': status
            })

    return render_template('index.html', vehicles=vehicles, busy_list=busy_list)

# --- 2. หน้าผู้อนุมัติ (Admin) ---
# [เพิ่ม] ระบบ Login ฝั่ง Admin
@app.route('/approve/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        password = request.form.get('password')
        if password == ADMIN_PASSWORD:
            session['logged_in'] = True
            flash('เข้าสู่ระบบสำเร็จ', 'success')
            return redirect(url_for('admin_approve'))
        else:
            flash('รหัสผ่านไม่ถูกต้อง', 'danger')
    return render_template('login.html')

# [เพิ่ม] ออกจากระบบ
@app.route('/approve/logout')
def admin_logout():
    session.pop('logged_in', None)
    flash('ออกจากระบบแล้ว', 'info')
    return redirect(url_for('admin_login'))

# [เปลี่ยน] ย้าย route จาก /admin เป็น /approve และบังคับ Login
@app.route('/approve')
@login_required
def admin_approve():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb['Bookings']
    
    bookings = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        bookings.append({
            'id': row[0], 'date': row[1], 'start': row[2], 'end': row[3],
            'vehicle': row[4], 'user': row[5], 'location': row[6],
            'purpose': row[7], 'passengers': row[8], 'note': row[9], 'status': row[10]
        })
    
    pending_bookings = [b for b in bookings if b['status'] == 'รออนุมัติ']
    return render_template('admin.html', pending_bookings=pending_bookings, all_bookings=bookings)

# [เปลี่ยน] route และบังคับ Login ก่อนเปลี่ยนสถานะ
@app.route('/approve/update_status/<booking_id>/<status>')
@login_required
def update_status(booking_id, status):
    # ป้องกันการยิง URL เข้ามาด้วย status แปลกๆ
    if status not in ['อนุมัติ', 'ไม่อนุมัติ']:
        flash('สถานะไม่ถูกต้อง', 'danger')
        return redirect(url_for('admin_approve'))

    with excel_lock: # [เพิ่ม] File lock
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb['Bookings']
        
        for row in ws.iter_rows(min_row=2):
            if row[0].value == booking_id:
                row[10].value = status
                break
                
        wb.save(EXCEL_FILE)
        
    log_action(f'อัปเดตสถานะเป็น {status}', booking_id, 'Admin')
    flash(f'ปรับสถานะ {booking_id} เป็น "{status}" เรียบร้อย', 'info')
    return redirect(url_for('admin_approve'))

# --- ตารางการใช้รถ (Calendar View) ---
@app.route('/calendar')
def calendar():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb['Bookings']
    
    events = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        events.append({
            'id': row[0], 'date': row[1], 'start': row[2], 'end': row[3],
            'vehicle': row[4], 'user': row[5], 'purpose': row[7], 'status': row[10]
        })
    return render_template('calendar.html', events=events)

# --- ตารางการใช้รถสำหรับแอดมิน (Admin Calendar View) ---
@app.route('/approve/calendar')
@login_required
def admin_calendar():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb['Bookings']
    
    events = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        events.append({
            'id': row[0], 'date': row[1], 'start': row[2], 'end': row[3],
            'vehicle': row[4], 'user': row[5], 'purpose': row[7], 'status': row[10]
        })
    return render_template('calendar.html', events=events)

# Redirect /admin to /approve
@app.route('/admin')
def admin_shortcut():
    return redirect(url_for('admin_approve'))

# --- หน้าผู้จัดทำ (Creators View) ---
MEMBERS_DATA = [
    {"name": "ธนาวดี ธะนะ", "id": "67040160", "role": "สมาชิกทีมพัฒนา"},
    {"name": "กชพรรณ ปิตุโส", "id": "68040001", "role": "สมาชิกทีมพัฒนา"},
    {"name": "จิรายุทธ โพธิ์ทอง", "id": "68040080", "role": "สมาชิกทีมพัฒนา"},
    {"name": "ธัญวรัตน์ กิจหงวน", "id": "68040239", "role": "สมาชิกทีมพัฒนา"},
    {"name": "นันท์นภัส ตั้งจิตร์", "id": "68040276", "role": "สมาชิกทีมพัฒนา"},
    {"name": "กาญจนา พวงพันธ์", "id": "68040617", "role": "สมาชิกทีมพัฒนา"},
    {"name": "ณัฐกฤตา ดวงจันทร์", "id": "68040686", "role": "สมาชิกทีมพัฒนา"},
]

@app.route('/creators')
def creators():
    return render_template('creators.html', members=MEMBERS_DATA)

@app.route('/approve/creators')
@login_required
def admin_creators():
    return render_template('creators.html', members=MEMBERS_DATA)

@app.route('/about')
def about_shortcut():
    return redirect(url_for('creators'))



if __name__ == '__main__':
    init_excel()
    app.run(host='0.0.0.0', port=8000, debug=True)