from datetime import datetime
from functools import wraps
import os
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
import gspread

app = Flask(__name__)

# ดึง secret_key และรหัสผ่านแอดมิน
app.secret_key = os.environ.get(
    'FLASK_SECRET_KEY', 'default_fallback_secret_key'
)
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')

# --- เชื่อมต่อ Google Sheets ---
gc = gspread.service_account(filename='credentials.json')
sh = gc.open('CarBooking')  # ชื่อ Google Sheets ใน Google Drive

# เชื่อมต่อแท็บแผ่นงาน
ws_bookings = sh.worksheet('Bookings')
ws_vehicles = sh.worksheet('Vehicles')
ws_users = sh.worksheet('Users')
ws_logs = sh.worksheet('Logs')


def init_excel():
  """ฟังก์ชันเดิมสำหรับสร้างไฟล์ Excel

  เมื่อเปลี่ยนไปใช้ Google Sheets แล้วไม่จำเป็นต้องสร้างไฟล์ในเครื่อง
  """
  pass


def log_action(action, detail, user):
  """บันทึกประวัติการกระทำลงในชีต Logs"""
  now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
  ws_logs.append_row([now_str, action, detail, user])


# Decorator สำหรับบังคับล็อกอินส่วนของ Admin
def login_required(f):

  @wraps(f)
  def decorated_function(*args, **kwargs):
    if not session.get('logged_in'):
      flash('กรุณาล็อกอินก่อนเข้าใช้งานระบบผู้อนุมัติ', 'warning')
      return redirect(url_for('admin_login'))
    return f(*args, **kwargs)

  return decorated_function


# --- 1. หน้าผู้จอง (User Form) ---
@app.route('/')
def index():
  return redirect(url_for('user_request'))


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

    # ดึงข้อมูลทั้งหมดจากชีต Bookings มาตรวจสอบ
    all_booking_rows = ws_bookings.get_all_values()

    # ตรวจสอบว่าช่วงเวลาดังกล่าวถูกจองไปแล้วหรือไม่ (ข้ามแถวหัวตาราง)
    if len(all_booking_rows) > 1:
      for row in all_booking_rows[1:]:
        if len(row) >= 11:
          ex_date = str(row[1])
          ex_start = str(row[2])
          ex_end = str(row[3])
          ex_vehicle = str(row[4])
          ex_status = str(row[10])

          if (
              ex_vehicle == vehicle
              and ex_date == date_use
              and ex_status != 'ไม่อนุมัติ'
          ):
            if start_time < ex_end and end_time > ex_start:
              flash(
                  f'ไม่สามารถจองได้: รถ "{vehicle}" ในวันที่ {date_use}'
                  f' ช่วงเวลา {ex_start} ถึง {ex_end} มีการจองแล้ว!',
                  'danger',
              )
              return redirect(url_for('user_request'))

    # คำนวณ Booking ID จากจำนวนแถวที่มีอยู่
    total_rows = len(all_booking_rows)
    booking_id = f'BK{total_rows:03d}'

    # บันทึกข้อมูลการจองแถวใหม่ลง Google Sheets
    ws_bookings.append_row([
        booking_id,
        date_use,
        start_time,
        end_time,
        vehicle,
        user_name,
        location,
        purpose,
        passengers,
        note,
        'รออนุมัติ',
    ])

    log_action('สร้างคำขอจอง', booking_id, user_name)
    flash('ส่งคำขอจองรถเรียบร้อยแล้ว!', 'success')
    return redirect(url_for('user_request'))

  # === ส่วนแสดงผลหน้าเว็บ (GET) ===
  vehicle_rows = ws_vehicles.get_all_values()
  # ดึงชื่อรถจากคอลัมน์ที่ 2 (index 1) ข้ามหัวตาราง
  vehicles = (
      [r[1] for r in vehicle_rows[1:] if len(r) > 1]
      if len(vehicle_rows) > 1
      else []
  )

  booking_rows = ws_bookings.get_all_values()
  busy_list = []
  if len(booking_rows) > 1:
    for row in booking_rows[1:]:
      if len(row) >= 11:
        status = row[10]
        if status in ['อนุมัติ', 'รออนุมัติ']:
          busy_list.append({
              'vehicle': row[4],
              'date': row[1],
              'time': f'{row[2]} - {row[3]}',
              'status': status,
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
  booking_rows = ws_bookings.get_all_values()

  bookings = []
  if len(booking_rows) > 1:
    for row in booking_rows[1:]:
      if len(row) >= 11:
        bookings.append({
            'id': row[0],
            'date': row[1],
            'start': row[2],
            'end': row[3],
            'vehicle': row[4],
            'user': row[5],
            'location': row[6],
            'purpose': row[7],
            'passengers': row[8],
            'note': row[9],
            'status': row[10],
        })

  pending_bookings = [b for b in bookings if b['status'] == 'รออนุมัติ']
  return render_template(
      'admin.html', pending_bookings=pending_bookings, all_bookings=bookings
  )


@app.route('/approve/update_status/<booking_id>/<status>')
@login_required
def update_status(booking_id, status):
  if status not in ['อนุมัติ', 'ไม่อนุมัติ']:
    flash('สถานะไม่ถูกต้อง', 'danger')
    return redirect(url_for('admin_approve'))

  booking_rows = ws_bookings.get_all_values()

  # ค้นหาแถวที่ตรงกับ Booking ID แล้วอัปเดตสถานะ (คอลัมน์ที่ 11 หรือ K)
  for index, row in enumerate(booking_rows[1:], start=2):
    if len(row) > 0 and row[0] == booking_id:
      ws_bookings.update_cell(index, 11, status)
      break

  log_action(f'อัปเดตสถานะเป็น {status}', booking_id, 'Admin')
  flash(f'ปรับสถานะ {booking_id} เป็น "{status}" เรียบร้อย', 'info')
  return redirect(url_for('admin_approve'))


# --- ตารางการใช้รถ (Calendar View) ---
@app.route('/calendar')
def calendar():
  booking_rows = ws_bookings.get_all_values()

  events = []
  if len(booking_rows) > 1:
    for row in booking_rows[1:]:
      if len(row) >= 11:
        events.append({
            'id': row[0],
            'date': row[1],
            'start': row[2],
            'end': row[3],
            'vehicle': row[4],
            'user': row[5],
            'purpose': row[7],
            'status': row[10],
        })
  return render_template('calendar.html', events=events)


@app.route('/approve/calendar')
@login_required
def admin_calendar():
  return calendar()


@app.route('/admin')
def admin_shortcut():
  return redirect(url_for('admin_approve'))


# --- หน้าผู้จัดทำ (Creators View) ---
MEMBERS_DATA = [
    {'name': 'ธนาวดี ธะนะ', 'id': '67040160', 'role': 'สมาชิกทีมพัฒนา'},
    {'name': 'กชพรรณ ปิตุโส', 'id': '68040001', 'role': 'สมาชิกทีมพัฒนา'},
    {'name': 'จิรายุทธ โพธิ์ทอง', 'id': '68040080', 'role': 'สมาชิกทีมพัฒนา'},
    {'name': 'ธัญวรัตน์ กิจหงวน', 'id': '68040239', 'role': 'สมาชิกทีมพัฒนา'},
    {'name': 'นันท์นภัส ตั้งจิตร์', 'id': '68040276', 'role': 'สมาชิกทีมพัฒนา'},
    {'name': 'กาญจนา พวงพันธ์', 'id': '68040617', 'role': 'สมาชิกทีมพัฒนา'},
    {'name': 'ณัฐกฤตา ดวงจันทร์', 'id': '68040686', 'role': 'สมาชิกทีมพัฒนา'},
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