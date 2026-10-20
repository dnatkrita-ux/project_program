import os
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, session
import openpyxl
from openpyxl import Workbook

try:
    from filelock import FileLock
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

app = Flask(__name__)

# ดึง secret_key และรหัสผ่านแอดมิน
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'default_fallback_secret_key')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'CarBooking.xlsx')
LOCK_FILE = os.path.join(BASE_DIR, 'CarBooking.xlsx.lock')
excel_lock = FileLock(LOCK_FILE, timeout=10)

# รายการรถและเครื่องจักรกลการเกษตรที่จำเป็นและนิยมใช้ในไทย (ครอบคลุม พืชไร่ นาข้าว พืชสวน ปาล์ม และยางพารา)
DEFAULT_VEHICLES = [
    ('V001', 'รถแทรกเตอร์เพื่อการเกษตร', 'ษ 8678'),
    ('V002', 'รถเกี่ยวนวดข้าว', 'นข 1234'),
    ('V003', 'รถดำนา', 'ฟข 8765'),
    ('V004', 'รถเกี่ยวอ้อย', 'นข 5678'),
    ('V005', 'รถขุดมันสำปะหลัง', 'กข 1111'),
    ('V006', 'รถพ่นยาเพื่อการเกษตร', 'ญณ 1291'),
    ('V007', 'รถอีแต๋นเพื่อการเกษตร', 'ตถ 9876'),
    ('V008', 'รถอัดฟางก้อน', 'อฟ 3456'),
    ('V009', 'รถหยอดเมล็ดพันธุ์', 'มพ 7890'),
    ('V010', 'รถบรรทุกผลผลิตการเกษตร', 'ปป 2233'),
    ('V011', 'รถไถเดินตาม (ควายเหล็ก)', 'คห 1011'),
    ('V012', 'รถคีบอ้อย / รถตักการเกษตร', 'ตอ 4455'),
    ('V013', 'รถขุดไฮดรอลิกเพื่อการเกษตร (มินิแม็คโครเกษตร)', 'มค 7788'),
    ('V014', 'รถอีแต๊กเพื่อการเกษตร', 'อต 3344'),
    ('V015', 'รถแทรกเตอร์สวนผลไม้ (ทรงแคบ)', 'สผ 5566'),
    ('V016', 'รถพ่นยาขับเคลื่อน 4 ล้อทรงสูง (High-clearance)', 'พส 8899'),
    ('V017', 'รถตัดหญ้าเพื่อการเกษตรแบบนั่งขับ', 'ตห 2244'),
    ('V018', 'รถขนส่งผลปาล์มน้ำมันและยางพารา', 'ปม 6677'),
    ('V019', 'รถหว่านปุ๋ยและปูนโดโลไมต์', 'หป 9911'),
    ('V020', 'รถเกลี่ยดินปรับระดับเลเซอร์แปลงนา', 'กด 1221'),
    ('V021', 'รถปลูกอ้อยและหยอดปุ๋ย', 'ปอ 3322'),
    ('V022', 'รถตัดหญ้าอาหารสัตว์และสับหญ้าเนเปียร์', 'สน 7711')
]

# --- เชื่อมต่อ Google Sheets (ถ้ามี credentials และต่อเน็ตได้) หากเชื่อมไม่ได้จะใช้ Local Excel ---
use_google_sheets = False
gc = None
sh = None
ws_bookings_g = None
ws_vehicles_g = None
ws_logs_g = None

cred_path = os.path.join(BASE_DIR, 'credentials.json')
if os.path.exists(cred_path):
    try:
        import gspread
        gc = gspread.service_account(filename=cred_path)
        sh = gc.open('CarBooking')
        ws_bookings_g = sh.worksheet('Bookings')
        ws_vehicles_g = sh.worksheet('Vehicles')
        ws_logs_g = sh.worksheet('Logs')
        use_google_sheets = True
        print("Connected to Google Sheets: CarBooking")
    except Exception as e:
        print(f"Notice: Google Sheets unavailable ({e}), running in local Excel mode.")
        use_google_sheets = False

def init_excel():
    """สร้างหรืออัปเดตไฟล์ Excel ในเครื่องให้มีตารางและรายชื่อรถการเกษตรครบถ้วน"""
    with excel_lock:
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            
            # Sheet 1: Bookings
            ws_bookings = wb.active
            ws_bookings.title = 'Bookings'
            ws_bookings.append([
                'รหัสการจอง', 'วันที่เริ่มต้น', 'เวลาเริ่มต้น', 'วันที่สิ้นสุด', 'เวลาสิ้นสุด', 
                'ยานพาหนะ', 'ผู้ขอใช้รถ', 'สถานที่ปฏิบัติงาน', 'วัตถุประสงค์การใช้งาน', 'หมายเหตุ', 'สถานะ'
            ])
            
            # ข้อมูลเริ่มต้น
            ws_bookings.append(['BK001', '2026-10-05', '08:30', '2026-10-05', '12:00', 'รถเกี่ยวนวดข้าว', 'นาย A', 'พื้นที่แปลงเกษตรตนเอง', 'เก็บเกี่ยวผลผลิตข้าว', '-', 'อนุมัติ'])
            ws_bookings.append(['BK002', '2026-10-05', '13:00', '2026-10-05', '16:30', 'รถเกี่ยวอ้อย', 'นาย B', 'ไร่อ้อยแปลงที่ 2', 'เก็บเกี่ยวอ้อยส่งโรงงาน', '-', 'อนุมัติ'])
            ws_bookings.append(['BK003', '2026-10-06', '09:00', '2026-10-06', '15:00', 'รถขุดมันสำปะหลัง', 'นาย C', 'แปลงมันสำปะหลัง ระยอง', 'เก็บเกี่ยวผลผลิตมันสำปะหลัง', '-', 'ไม่อนุมัติ'])
            
            # Sheet 2: Vehicles
            ws_vehicles = wb.create_sheet(title='Vehicles')
            ws_vehicles.append(['รหัสยานพาหนะ', 'ประเภทยานพาหนะ', 'ทะเบียน'])
            for v in DEFAULT_VEHICLES:
                ws_vehicles.append(list(v))
                
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
        else:
            # อัปเดตรถใหม่เข้า Sheet Vehicles หากยังไม่มี
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if 'Vehicles' in wb.sheetnames:
                ws_vehicles = wb['Vehicles']
                existing_codes = set()
                for r in ws_vehicles.iter_rows(min_row=2, values_only=True):
                    if r and r[0]:
                        existing_codes.add(str(r[0]).strip())
                
                updated = False
                for v in DEFAULT_VEHICLES:
                    if v[0] not in existing_codes:
                        ws_vehicles.append(list(v))
                        updated = True
                if updated:
                    wb.save(EXCEL_FILE)

def parse_booking_row(row):
    """ฟังก์ชันแปลงแถวข้อมูลการจองให้เป็น Dictionary อย่างปลอดภัย"""
    if not row or not row[0]:
        return None
    
    # รองรับทั้งสกีมา 11 คอลัมน์ (มี start_date, start_time, end_date, end_time) 
    # และสกีมาเดิม (date, start, end, vehicle...)
    if len(row) >= 11:
        # สกีมาใหม่ 11 คอลัมน์: รหัส, วันเริ่ม, เวลาเริ่ม, วันสิ้นสุด, เวลาสิ้นสุด, รถ, ผู้ขอ, สถานที่, วัตถุประสงค์, หมายเหตุ, สถานะ
        return {
            'id': str(row[0]),
            'start_date': str(row[1]),
            'start_time': str(row[2]),
            'end_date': str(row[3]),
            'end_time': str(row[4]),
            'vehicle': str(row[5]),
            'user': str(row[6]),
            'location': str(row[7]),
            'purpose': str(row[8]),
            'note': str(row[9]) if (row[9] is not None and str(row[9]).strip() != '') else '-',
            'status': str(row[10]) if row[10] is not None else 'รออนุมัติ'
        }
    return None

def get_all_bookings_data():
    """ดึงข้อมูลรายการจองทั้งหมดจากแหล่งข้อมูล"""
    bookings = []
    if use_google_sheets and ws_bookings_g:
        try:
            rows = ws_bookings_g.get_all_values()
            for r in rows[1:]:
                item = parse_booking_row(r)
                if item:
                    bookings.append(item)
            return bookings
        except Exception as e:
            print(f"Error reading from Google Sheets: {e}")
            
    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb['Bookings']
        for row in ws.iter_rows(min_row=2, values_only=True):
            item = parse_booking_row(row)
            if item:
                bookings.append(item)
    return bookings

def get_vehicles_list():
    """ดึงรายชื่อรถและเครื่องจักรกลการเกษตรสำหรับ Dropdown"""
    vehicles = []
    if use_google_sheets and ws_vehicles_g:
        try:
            rows = ws_vehicles_g.get_all_values()
            for r in rows[1:]:
                if len(r) > 1 and r[1]:
                    plate = f" ({r[2]})" if len(r) > 2 and r[2] else ""
                    vehicles.append(f"{r[1]}{plate}")
            if vehicles:
                return vehicles
        except Exception as e:
            print(f"Error reading vehicles from Google Sheets: {e}")

    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb['Vehicles']
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and len(r) > 1 and r[1]:
                plate = f" ({r[2]})" if len(r) > 2 and r[2] else ""
                vehicles.append(f"{r[1]}{plate}")
    return vehicles or [f"{v[1]} ({v[2]})" for v in DEFAULT_VEHICLES]

def log_action(action, detail, user):
    """บันทึกประวัติการกระทำลงใน Logs"""
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if use_google_sheets and ws_logs_g:
        try:
            ws_logs_g.append_row([now_str, action, detail, user])
            return
        except Exception as e:
            print(f"Error logging to Google Sheets: {e}")

    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_logs = wb['Logs']
        ws_logs.append([now_str, action, detail, user])
        wb.save(EXCEL_FILE)

# Decorator สำหรับบังคับล็อกอินส่วนของ Admin
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in'):
            flash('กรุณาล็อกอินก่อนเข้าใช้งานระบบผู้อนุมัติ', 'warning')
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    return decorated_function

# --- 1. หน้าผู้ขอใช้รถ (User Form) ---
@app.route('/')
def index():
    return redirect(url_for('user_request'))

@app.route('/request', methods=['GET', 'POST'])
def user_request():
    if request.method == 'POST':
        user_name = request.form.get('user_name', '').strip()
        start_date = (request.form.get('start_date') or request.form.get('date_use') or '').strip()
        start_time = request.form.get('start_time', '').strip()
        end_date = (request.form.get('end_date') or start_date or '').strip()
        end_time = request.form.get('end_time', '').strip()
        vehicle = request.form.get('vehicle', '').strip()
        location = request.form.get('location', '').strip()
        purpose = request.form.get('purpose', '').strip()
        note = request.form.get('note', '-').strip() or '-'

        # ตรวจสอบข้อมูลจำเป็น
        if not all([user_name, start_date, start_time, end_date, end_time, vehicle, location, purpose]):
            flash('กรุณากรอกข้อมูลในช่องที่มีเครื่องหมายดอกจัน (*) ให้ครบถ้วน', 'warning')
            return redirect(url_for('user_request'))

        # ตรวจสอบว่าวันสิ้นสุดต้องไม่มาก่อนวันเริ่มต้น
        req_start = f"{start_date} {start_time}"
        req_end = f"{end_date} {end_time}"
        if req_end <= req_start:
            flash('วันและเวลาสิ้นสุดต้องอยู่หลังจากวันและเวลาเริ่มต้นการใช้งาน', 'danger')
            return redirect(url_for('user_request'))

        # ตรวจสอบการชนกันของช่วงเวลาการจอง (Overlap Check)
        all_bookings = get_all_bookings_data()
        for b in all_bookings:
            if b['status'] != 'ไม่อนุมัติ':
                # ตรวจสอบชื่อรถ (จับคู่ทั้งแบบมีทะเบียนและไม่มีทะเบียน)
                v_booked = b['vehicle'].strip()
                v_request = vehicle.strip()
                is_same_vehicle = (
                    v_booked == v_request or 
                    v_booked in v_request or 
                    v_request in v_booked
                )
                if is_same_vehicle:
                    b_start = f"{b['start_date']} {b['start_time']}"
                    b_end = f"{b['end_date']} {b['end_time']}"
                    # เงื่อนไขการทับซ้อน: req_start < b_end and req_end > b_start
                    if req_start < b_end and req_end > b_start:
                        flash(
                            f'ไม่สามารถจองได้: ยานพาหนะ "{vehicle}" ถูกจองแล้วในช่วงวันที่ '
                            f'{b["start_date"]} ({b["start_time"]}) ถึง {b["end_date"]} ({b["end_time"]})!',
                            'danger'
                        )
                        return redirect(url_for('user_request'))

        # กำหนดรหัสการจองใหม่
        booking_id = f"BK{len(all_bookings) + 1:03d}"
        new_row = [
            booking_id, start_date, start_time, end_date, end_time,
            vehicle, user_name, location, purpose, note, 'รออนุมัติ'
        ]

        # บันทึกลงระบบ
        saved = False
        if use_google_sheets and ws_bookings_g:
            try:
                ws_bookings_g.append_row(new_row)
                saved = True
            except Exception as e:
                print(f"Error saving to Google Sheets: {e}")

        # บันทึกลง Local Excel เสมอเพื่อความเสถียรและความปลอดภัย
        with excel_lock:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws_bookings = wb['Bookings']
            ws_bookings.append(new_row)
            wb.save(EXCEL_FILE)
            saved = True

        log_action('สร้างคำขอจอง', booking_id, user_name)
        flash('ส่งคำขอจองยานพาหนะการเกษตรเรียบร้อยแล้ว!', 'success')
        return redirect(url_for('user_request'))

    # === ส่วนแสดงผลหน้าเว็บ (GET) ===
    vehicles = get_vehicles_list()
    all_bookings = get_all_bookings_data()

    # จัดเตรียมข้อมูลตารางช่วงเวลาที่ไม่ว่าง (Real-time Availability)
    busy_list = []
    for b in all_bookings:
        if b['status'] in ['อนุมัติ', 'รออนุมัติ']:
            date_display = b['start_date'] if b['start_date'] == b['end_date'] else f"{b['start_date']} ถึง {b['end_date']}"
            time_display = f"{b['start_time']} - {b['end_time']}"
            busy_list.append({
                'vehicle': b['vehicle'],
                'date_display': date_display,
                'time_display': time_display,
                'status': b['status']
            })

    return render_template('index.html', vehicles=vehicles, busy_list=busy_list)

# --- 2. หน้าผู้อนุมัติ (Admin) ---
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

@app.route('/approve/logout')
def admin_logout():
    session.pop('logged_in', None)
    flash('ออกจากระบบแล้ว', 'info')
    return redirect(url_for('admin_login'))

@app.route('/approve')
@login_required
def admin_approve():
    all_bookings = get_all_bookings_data()
    pending_bookings = [b for b in all_bookings if b['status'] == 'รออนุมัติ']
    return render_template(
        'admin.html',
        pending_bookings=pending_bookings,
        all_bookings=all_bookings
    )

@app.route('/approve/update_status/<booking_id>/<status>')
@login_required
def update_status(booking_id, status):
    if status not in ['อนุมัติ', 'ไม่อนุมัติ']:
        flash('สถานะไม่ถูกต้อง', 'danger')
        return redirect(url_for('admin_approve'))

    # อัปเดตใน Google Sheets
    if use_google_sheets and ws_bookings_g:
        try:
            rows = ws_bookings_g.get_all_values()
            for index, row in enumerate(rows[1:], start=2):
                if len(row) > 0 and row[0] == booking_id:
                    ws_bookings_g.update_cell(index, 11, status)
                    break
        except Exception as e:
            print(f"Error updating Google Sheets: {e}")

    # อัปเดตใน Local Excel
    with excel_lock:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb['Bookings']
        for row in ws.iter_rows(min_row=2):
            if row[0].value == booking_id:
                row[10].value = status
                break
        wb.save(EXCEL_FILE)

    log_action(f'อัปเดตสถานะเป็น {status}', booking_id, 'Admin')
    flash(f'ปรับสถานะคำขอ {booking_id} เป็น "{status}" เรียบร้อยแล้ว', 'info')
    return redirect(url_for('admin_approve'))

# --- 3. หน้าตารางการใช้รถ (Schedule / Calendar View) ---
@app.route('/calendar')
def calendar():
    all_bookings = get_all_bookings_data()
    events = []
    for b in all_bookings:
        events.append({
            'id': b['id'],
            'start_date': b['start_date'],
            'start_time': b['start_time'],
            'end_date': b['end_date'],
            'end_time': b['end_time'],
            'vehicle': b['vehicle'],
            'user': b['user'],
            'purpose': b['purpose'],
            'status': b['status']
        })
    return render_template('calendar.html', events=events)

@app.route('/approve/calendar')
@login_required
def admin_calendar():
    return calendar()

@app.route('/admin')
def admin_shortcut():
    return redirect(url_for('admin_approve'))

# --- 4. หน้าคณะผู้จัดทำ (Creators View) ---
MEMBERS_DATA = [
    {"name": "ธนาวดี ธะนะ", "id": "67040160", "role": "สมาชิกทีมพัฒนา"},
    {"name": "กชพรรณ ปิตุโส", "id": "68040001", "role": "สมาชิกทีมพัฒนา"},
    {"name": "จิรายุทธ โพธิ์ทอง", "id": "68040080", "role": "สมาชิกทีมพัฒนา"},
    {"name": "ธัญวรัตน์ กิจหงวน", "id": "68040239", "role": "สมาชิกทีมพัฒนา"},
    {"name": "นันท์นภัส ตั้งจิตร์", "id": "68040276", "role": "สมาชิกทีมพัฒนา"},
    {"name": "กาญจนา พวงพันธ์", "id": "68040617", "role": "สมาชิกทีมพัฒนา"},
    {"name": "ณัฐกฤตา ดวงจันทร์", "id": "68040686", "role": "สมาชิกทีมพัฒนา"}
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