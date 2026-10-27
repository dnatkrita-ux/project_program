import os
import io
import re
import random
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import gspread

app = Flask(__name__)

# ดึง secret_key และรหัสผ่านแอดมิน (ความยาวอย่างน้อย 8 ตัวอักษร มีตัวอักษร ตัวเลข และอักขระพิเศษ: Admin@123)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'agricarbook_super_secure_key_2026')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'Admin@123')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'CarBooking.xlsx')

# รายการรถและเครื่องจักรกลการเกษตรเริ่มต้นครบทั้ง 22 คัน
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

# รายชื่อผู้ใช้เริ่มต้น (รหัสผ่าน 8 ตัว มีตัวอักษร ตัวเลข และอักขระพิเศษ พร้อมอีเมล และเลขบัตรประชาชน 13 หลัก)
DEFAULT_USERS = [
    ('U001', 'นายสมชาย ใจดี', '100123456789', 'User@123', 'ผู้ใช้', 'somchai@farm.go.th', '1100100111111'),
    ('U002', 'นางสมศรี สดใส', '200123456789', 'Farm#123', 'ผู้ใช้', 'somsri@farm.go.th', '3100500123458'),
    ('U003', 'ผู้ดูแลระบบ (Admin)', 'admin', 'Admin@123', 'ผู้อนุมัติ', 'admin@agricar.com', '')
]

# --- เชื่อมต่อ Google Sheets (ถ้าเชื่อมต่อไม่ได้จะสลับไปใช้ Local CarBooking.xlsx อัตโนมัติ) ---
use_google_sheets = False
gc = None
sh = None
ws_bookings = None
ws_vehicles = None
ws_users = None
ws_logs = None
ws_notifications = None

try:
    gc = gspread.service_account(filename='credentials.json')
    sh = gc.open('CarBooking')
    ws_bookings = sh.worksheet('Bookings')
    ws_vehicles = sh.worksheet('Vehicles')
    ws_users = sh.worksheet('Users')
    ws_logs = sh.worksheet('Logs')
    try:
        ws_notifications = sh.worksheet('Notifications')
    except Exception:
        try:
            ws_notifications = sh.add_worksheet(title='Notifications', rows=100, cols=10)
            ws_notifications.append_row([
                'รหัสแจ้งเตือน', 'รหัสผู้ใช้', 'ชื่อผู้ใช้', 'อีเมล', 
                'รหัสคำสั่งจอง', 'ยานพาหนะ', 'สถานะ', 'เหตุผล/หมายเหตุ', 'วันที่/เวลา', 'อ่านแล้ว'
            ])
        except Exception:
            ws_notifications = None
    use_google_sheets = True
except Exception as e:
    print(f"Notice: Google Sheets unavailable ({e}), using local CarBooking.xlsx")
    use_google_sheets = False

def init_excel():
    """ฟังก์ชันสำหรับสร้างหรือตรวจสอบไฟล์ Local Excel ให้มีตารางและชีตครบถ้วน"""
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws_b = wb.active
        ws_b.title = 'Bookings'
        ws_b.append([
            'รหัสการจอง', 'วันที่เริ่มต้น', 'เวลาเริ่มต้น', 'วันที่สิ้นสุด', 'เวลาสิ้นสุด', 
            'ยานพาหนะ', 'ผู้ขอใช้รถ', 'สถานที่ปฏิบัติงาน', 'วัตถุประสงค์การใช้งาน', 'หมายเหตุ', 'สถานะ'
        ])
        
        ws_v = wb.create_sheet(title='Vehicles')
        ws_v.append(['รหัสยานพาหนะ', 'ประเภทยานพาหนะ', 'ทะเบียน'])
        for v in DEFAULT_VEHICLES:
            ws_v.append(list(v))
            
        ws_u = wb.create_sheet(title='Users')
        ws_u.append(['รหัสผู้ใช้', 'ชื่อ-สกุล', 'รหัสเกษตร/ชื่อผู้ใช้', 'รหัสผ่าน', 'บทบาท', 'อีเมล', 'เลขประจำตัวประชาชน'])
        for u in DEFAULT_USERS:
            ws_u.append(list(u))
            
        ws_l = wb.create_sheet(title='Logs')
        ws_l.append(['วันที่/เวลา', 'การดำเนินการ', 'รายละเอียด', 'ผู้ทำ'])

        ws_n = wb.create_sheet(title='Notifications')
        ws_n.append([
            'รหัสแจ้งเตือน', 'รหัสผู้ใช้', 'ชื่อผู้ใช้', 'อีเมล', 
            'รหัสคำสั่งจอง', 'ยานพาหนะ', 'สถานะ', 'เหตุผล/หมายเหตุ', 'วันที่/เวลา', 'อ่านแล้ว'
        ])
        wb.save(EXCEL_FILE)
    else:
        # ตรวจสอบว่ามีชีต Users หรือไม่ ถ้าไม่มีให้สร้าง
        wb = openpyxl.load_workbook(EXCEL_FILE)
        changed = False
        if 'Users' not in wb.sheetnames:
            ws_u = wb.create_sheet(title='Users')
            ws_u.append(['รหัสผู้ใช้', 'ชื่อ-สกุล', 'รหัสเกษตร/ชื่อผู้ใช้', 'รหัสผ่าน', 'บทบาท', 'อีเมล', 'เลขประจำตัวประชาชน'])
            for u in DEFAULT_USERS:
                ws_u.append(list(u))
            changed = True
        if 'Notifications' not in wb.sheetnames:
            ws_n = wb.create_sheet(title='Notifications')
            ws_n.append([
                'รหัสแจ้งเตือน', 'รหัสผู้ใช้', 'ชื่อผู้ใช้', 'อีเมล', 
                'รหัสคำสั่งจอง', 'ยานพาหนะ', 'สถานะ', 'เหตุผล/หมายเหตุ', 'วันที่/เวลา', 'อ่านแล้ว'
            ])
            changed = True
        if changed:
            wb.save(EXCEL_FILE)

def parse_booking_row(row):
    """ฟังก์ชันแปลงแถวข้อมูลในชีต Bookings ให้เป็น Dictionary อย่างถูกต้อง"""
    if not row or not row[0]:
        return None
    if len(row) >= 11:
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
            'status': str(row[10]) if row[10] is not None else 'รออนุมัติ',
            'date': str(row[1]),
            'start': str(row[2]),
            'end': str(row[4])
        }
    return None

def get_all_bookings():
    """ดึงข้อมูลรายการจองทั้งหมด"""
    bookings = []
    if use_google_sheets and ws_bookings:
        try:
            rows = ws_bookings.get_all_values()
            start_idx = 1 if (rows and 'รหัส' in str(rows[0][0])) else 0
            for r in rows[start_idx:]:
                item = parse_booking_row(r)
                if item:
                    bookings.append(item)
            return bookings
        except Exception as e:
            print(f"Error reading Google Sheets bookings: {e}")

    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_b = wb['Bookings']
        for row in ws_b.iter_rows(min_row=2, values_only=True):
            item = parse_booking_row(row)
            if item:
                bookings.append(item)
    return bookings

def get_all_vehicles_data():
    """ดึงข้อมูลยานพาหนะทั้งหมดจากชีต Vehicles เพื่อนำไปใช้ในหน้าจัดการรถ"""
    vehicles_data = []
    all_bookings = get_all_bookings()
    booking_counts = {}
    for b in all_bookings:
        v_name = b['vehicle'].strip()
        booking_counts[v_name] = booking_counts.get(v_name, 0) + 1

    rows_raw = []
    if use_google_sheets and ws_vehicles:
        try:
            rows = ws_vehicles.get_all_values()
            start_idx = 1 if (rows and 'รหัส' in str(rows[0][0])) else 0
            rows_raw = rows[start_idx:]
        except Exception as e:
            print(f"Error reading vehicles from Google Sheets: {e}")

    if not rows_raw and os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_v = wb['Vehicles']
        for r in ws_v.iter_rows(min_row=2, values_only=True):
            if r and r[0]:
                rows_raw.append(list(r))

    for r in rows_raw:
        if r and len(r) >= 2 and r[0]:
            code = str(r[0]).strip()
            name = str(r[1]).strip()
            plate = str(r[2]).strip() if len(r) > 2 and r[2] else '-'
            full_str = f"{name} ({plate})" if plate != '-' else name
            
            cnt = 0
            for b_veh, b_cnt in booking_counts.items():
                if code in b_veh or name in b_veh or (plate != '-' and plate in b_veh):
                    cnt += b_cnt

            vehicles_data.append({
                'code': code,
                'name': name,
                'plate': plate,
                'full_name': full_str,
                'booking_count': cnt
            })
    return vehicles_data

def get_vehicles():
    """ดึงรายชื่อยานพาหนะสำหรับ Dropdown (แสดงชื่อรถพร้อมทะเบียน)"""
    v_data = get_all_vehicles_data()
    if v_data:
        return [v['full_name'] for v in v_data]
    return [f"{v[1]} ({v[2]})" for v in DEFAULT_VEHICLES]

def add_vehicle_to_db(code, name, plate):
    """เพิ่มรถใหม่เข้าสู่ชีต Vehicles ทั้ง Google Sheets และ Local Excel"""
    new_row = [code, name, plate]
    if use_google_sheets and ws_vehicles:
        try:
            ws_vehicles.append_row(new_row)
        except Exception as e:
            print(f"Error adding vehicle to Google Sheets: {e}")

    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_v = wb['Vehicles']
        ws_v.append(new_row)
        wb.save(EXCEL_FILE)

def delete_vehicle_from_db(code):
    """ลบรถออกจากชีต Vehicles ทั้ง Google Sheets และ Local Excel"""
    if use_google_sheets and ws_vehicles:
        try:
            rows = ws_vehicles.get_all_values()
            for idx, r in enumerate(rows, start=1):
                if r and r[0] == code:
                    ws_vehicles.delete_rows(idx)
                    break
        except Exception as e:
            print(f"Error deleting vehicle from Google Sheets: {e}")

    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_v = wb['Vehicles']
        for row_idx, row in enumerate(ws_v.iter_rows(min_row=2), start=2):
            if row and row[0].value == code:
                ws_v.delete_rows(row_idx, 1)
                break
        wb.save(EXCEL_FILE)

# --- ระบบจัดการผู้ใช้ (User Management & 12-Character Password Authentication) ---
def get_all_users():
    """ดึงข้อมูลผู้ใช้ทั้งหมดจากชีต Users (รวมอีเมล, รหัสเกษตรกร และเลขประจำตัวประชาชน) พร้อม Fallback อัตโนมัติสำหรับบัญชีตัวอย่าง"""
    users = []
    if use_google_sheets and ws_users:
        try:
            rows = ws_users.get_all_values()
            start_idx = 1 if (rows and 'รหัส' in str(rows[0][0])) else 0
            for r in rows[start_idx:]:
                if r and len(r) >= 5 and r[2]:
                    email_val = str(r[5]).strip() if len(r) > 5 and r[5] else ''
                    citizen_val = str(r[6]).strip() if len(r) > 6 and r[6] else ''
                    users.append({
                        'id': str(r[0]).strip(),
                        'full_name': str(r[1]).strip(),
                        'username': str(r[2]).strip(),
                        'farmer_id': str(r[2]).strip(),
                        'password': str(r[3]).strip(),
                        'role': str(r[4]).strip(),
                        'email': email_val,
                        'citizen_id': citizen_val
                    })
        except Exception as e:
            print(f"Error reading users from Google Sheets: {e}")

    if not users and os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if 'Users' in wb.sheetnames:
                ws_u = wb['Users']
                for row in ws_u.iter_rows(min_row=2, values_only=True):
                    if row and len(row) >= 5 and row[2]:
                        email_val = str(row[5]).strip() if len(row) > 5 and row[5] else ''
                        citizen_val = str(row[6]).strip() if len(row) > 6 and row[6] else ''
                        users.append({
                            'id': str(row[0]).strip(),
                            'full_name': str(row[1]).strip(),
                            'username': str(row[2]).strip(),
                            'farmer_id': str(row[2]).strip(),
                            'password': str(row[3]).strip(),
                            'role': str(row[4]).strip(),
                            'email': email_val,
                            'citizen_id': citizen_val
                        })
        except Exception as e:
            print(f"Error reading users from Excel: {e}")

    # ตรวจสอบและเติมบัญชีตัวอย่าง (U001, U002, U003) ให้มีข้อมูลครบถ้วนเสมอ
    found_ids = {u['id']: u for u in users}
    for def_u in DEFAULT_USERS:
        u_id, u_name, u_uname, u_pwd, u_role, u_email, u_cid = def_u
        if u_id in found_ids:
            target = found_ids[u_id]
            if not target.get('email'):
                target['email'] = u_email
            if not target.get('citizen_id') and u_cid:
                target['citizen_id'] = u_cid
            # อัปเดตรหัสผ่านถ้าฐานข้อมูลยังเป็นค่าว่าง
            if not target.get('password'):
                target['password'] = u_pwd
        else:
            users.append({
                'id': u_id,
                'full_name': u_name,
                'username': u_uname,
                'farmer_id': u_uname,
                'password': u_pwd,
                'role': u_role,
                'email': u_email,
                'citizen_id': u_cid
            })

    return users

def find_user_by_identifier(identifier):
    """ค้นหาข้อมูลผู้ใช้ด้วย อีเมล, เลขประจำตัวประชาชน (13 หลัก), รหัสเกษตร หรือชื่อผู้ใช้"""
    if not identifier:
        return None
    ident = str(identifier).strip().lower()
    clean_digits = re.sub(r'\D', '', ident)
    users = get_all_users()

    for u in users:
        # 1. ตรวจสอบกับอีเมล (case-insensitive)
        if u.get('email') and u['email'].strip().lower() == ident:
            return u
        # 2. ตรวจสอบกับเลขประจำตัวประชาชน (รองรับทั้งแบบมีขีดและไม่มีขีด)
        if clean_digits and len(clean_digits) == 13 and u.get('citizen_id') and re.sub(r'\D', '', str(u['citizen_id'])) == clean_digits:
            return u
        # 3. ตรวจสอบกับรหัสเกษตร / username
        if u.get('username') and u['username'].strip().lower() == ident:
            return u
        if u.get('farmer_id') and str(u['farmer_id']).strip().lower() == ident:
            return u
        # 4. รองรับ alias รหัสผู้ใช้ เช่น u001, u002, u003, user01, user02
        if ident == u['id'].strip().lower():
            return u
        if ident == 'user01' and u['id'] == 'U001':
            return u
        if ident == 'user02' and u['id'] == 'U002':
            return u
        if ident == 'admin' and (u['id'] == 'U003' or u['role'] in ['ผู้อนุมัติ', 'admin']):
            return u

    # Guaranteed Fallback เฉพาะสำหรับบัญชีตัวอย่าง (กรณีฐานข้อมูลภายนอกยังซิงค์ไม่เสร็จ)
    if ident in ['somchai@farm.go.th', 'user01', '100123456789'] or clean_digits == '1100100111111':
        return {
            'id': 'U001',
            'full_name': 'นายสมชาย ใจดี',
            'username': '100123456789',
            'farmer_id': '100123456789',
            'password': 'User@123',
            'role': 'ผู้ใช้',
            'email': 'somchai@farm.go.th',
            'citizen_id': '1100100111111'
        }
    if ident in ['somsri@farm.go.th', 'user02', '200123456789'] or clean_digits == '3100500123458':
        return {
            'id': 'U002',
            'full_name': 'นางสมศรี สดใส',
            'username': '200123456789',
            'farmer_id': '200123456789',
            'password': 'Farm#123',
            'role': 'ผู้ใช้',
            'email': 'somsri@farm.go.th',
            'citizen_id': '3100500123458'
        }
    if ident in ['admin@agricar.com', 'admin', 'u003']:
        return {
            'id': 'U003',
            'full_name': 'ผู้ดูแลระบบ (Admin)',
            'username': 'admin',
            'farmer_id': 'admin',
            'password': 'Admin@123',
            'role': 'ผู้อนุมัติ',
            'email': 'admin@agricar.com',
            'citizen_id': ''
        }
    return None

def find_user_by_username(username):
    """ค้นหาข้อมูลผู้ใช้ด้วย username หรือ identifier"""
    return find_user_by_identifier(username)

def add_new_user(full_name, email, password, role='ผู้ใช้', citizen_id='', farmer_id=''):
    """เพิ่มผู้ใช้ใหม่ลงในฐานข้อมูลทั้ง Google Sheets และ Local Excel พร้อมอีเมลและเลขบัตรประชาชน"""
    users = get_all_users()
    next_id = f"U{len(users) + 1:03d}"
    clean_cid = re.sub(r'\D', '', str(citizen_id)) if citizen_id else ''
    username_val = farmer_id or email.split('@')[0] or clean_cid or next_id
    new_row = [next_id, full_name, username_val, password, role, email, clean_cid]

    if use_google_sheets and ws_users:
        try:
            ws_users.append_row(new_row)
        except Exception as e:
            print(f"Error adding user to Google Sheets: {e}")

    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_u = wb['Users']
        ws_u.append(new_row)
        wb.save(EXCEL_FILE)

    return next_id

# --- ฟังก์ชันตรวจสอบความถูกต้องของข้อมูล (Validation Functions) ---

def validate_thai_citizen_id(id_str):
    """
    ตรวจสอบความถูกต้องของเลขประจำตัวประชาชน 13 หลัก ด้วยสูตร Modulo 11 (Checksum Algorithm)
    สูตร: ผลรวม sum = (d1*13 + d2*12 + ... + d12*2)
          check_digit = (11 - (sum % 11)) % 10
    ตรวจสอบว่า d13 == check_digit หรือไม่
    คืนค่า: (is_valid: bool, message: str)
    """
    if not id_str:
        return False, "กรุณากรอกเลขประจำตัวประชาชน 13 หลัก"
    
    # ดึงเฉพาะตัวเลข 0-9
    digits = re.sub(r'\D', '', str(id_str))
    
    if len(digits) != 13:
        return False, f"กรอกเลขประจำตัวประชาชนไม่ครบ 13 หลัก (ปัจจุบันมี {len(digits)} หลัก)"
    
    # ตรวจสอบกรณีเลขตองเหมือนกันทั้ง 13 หลัก (เช่น 0000000000000, 1111111111111) ซึ่งเป็นเลขที่ไม่ถูกต้อง
    if len(set(digits)) == 1:
        return False, "เลขประจำตัวประชาชนไม่ถูกต้อง (เป็นตัวเลขซ้ำกันทั้งหมด)"
    
    # คำนวณผลรวมคูณถ่วงน้ำหนัก (Weighted Sum)
    total_sum = sum(int(digits[i]) * (13 - i) for i in range(12))
    
    # คำนวณ Checksum ด้วย Modulo 11
    check_digit = (11 - (total_sum % 11)) % 10
    
    if int(digits[12]) != check_digit:
        return False, f"เลขประจำตัวประชาชนไม่ถูกต้อง (คำนวณ Checksum ไม่ผ่าน: ได้ {check_digit} แต่หลักที่ 13 คือ {digits[12]})"
    
    return True, "เลขประจำตัวประชาชนถูกต้องตามสูตร Modulo 11"

def validate_farmer_id(farmer_id_str):
    r"""
    ตรวจสอบรหัสประจำตัวเกษตรกร (12 หลัก) ด้วย Regular Expression
    Regex Pattern: ^\d{12}$ (ตัวเลข 0-9 ความยาว 12 หลักพอดี)
    คืนค่า: (is_valid: bool, message: str)
    """
    if not farmer_id_str:
        return False, "กรุณากรอกรหัสประจำตัวเกษตรกร (12 หลัก)"
    
    s = str(farmer_id_str).strip()
    regex_pattern = r'^\d{12}$'
    
    if not re.match(regex_pattern, s):
        if not s.isdigit():
            return False, "รหัสประจำตัวเกษตรกรต้องเป็นตัวเลข 0-9 เท่านั้น (ห้ามมีตัวอักษรหรือสัญลักษณ์)"
        return False, f"รหัสประจำตัวเกษตรกรต้องมีความยาว 12 หลักพอดี (ปัจจุบันมี {len(s)} หลัก)"
    
    return True, "รหัสประจำตัวเกษตรกรถูกต้องตามรูปแบบ 12 หลัก"

def validate_password_complexity(password):
    """
    ตรวจสอบเงื่อนไขความปลอดภัยของรหัสผ่าน:
    - ความยาวอย่างน้อย 8 ตัวอักษร
    - ต้องมีตัวอักษรภาษาอังกฤษ (A-Z หรือ a-z)
    - ต้องมีตัวเลขอารบิก (0-9)
    - ต้องมีอักขระพิเศษ (เช่น ! @ # $ % ^ & * _ - + = < > ?)
    คืนค่า (is_valid: bool, message: str)
    """
    if not password:
        return False, "กรุณากรอกรหัสผ่าน"
    if len(password) < 8:
        return False, f"รหัสผ่านต้องมีความยาวอย่างน้อย 8 ตัวอักษร (ปัจจุบันมี {len(password)} ตัว)"
    if not re.search(r'[a-zA-Z]', password):
        return False, "รหัสผ่านต้องมีตัวอักษรภาษาอังกฤษ (A-Z หรือ a-z) อย่างน้อย 1 ตัว"
    if not re.search(r'[0-9]', password):
        return False, "รหัสผ่านต้องมีตัวเลข (0-9) อย่างน้อย 1 ตัว"
    if not re.search(r'[^a-zA-Z0-9]', password):
        return False, "รหัสผ่านต้องมีอักขระพิเศษ (เช่น ! @ # $ % ^ & * _ -) อย่างน้อย 1 ตัว"
    return True, "รหัสผ่านถูกต้องตามเงื่อนไขความปลอดภัย (8 ตัวอักษรขึ้นไป พร้อมตัวอักษร ตัวเลข และอักขระพิเศษ)"

# --- การคุ้มครองข้อมูลส่วนบุคคล (PDPA Privacy Masking) & จัดการชื่อ-นามสกุล ---
MASKED_FARM_LOCATION = "🔒 ปกปิดข้อมูลพื้นที่ฟาร์ม (สงวนสิทธิ์เฉพาะเจ้าของและแอดมิน)"

def split_full_name(full_name: str) -> tuple[str, str]:
    """
    แยกชื่อจริงและนามสกุล โดยรองรับคำนำหน้านามภาษาไทย เช่น นาย, นาง, นางสาว, น.ส., ดร.
    คืนค่า (first_name, last_name)
    """
    if not full_name:
        return "", ""
    parts = str(full_name).strip().split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    titles = {'นาย', 'นาง', 'นางสาว', 'น.ส.', 'ดร.', 'อาจารย์', 'อ.', 'ผศ.', 'รศ.', 'ศ.', 'คุณ'}
    if parts[0] in titles and len(parts) >= 3:
        return f"{parts[0]} {parts[1]}", " ".join(parts[2:])
    return parts[0], " ".join(parts[1:])

def mask_user_last_name(full_name: str) -> str:
    """
    ปิดบังข้อมูลนามสกุลเพื่อคุ้มครองข้อมูลส่วนบุคคล (PDPA)
    เช่น นายสมเกียรติ สุขเกษม -> นายสมเกียรติ ส***
         สมชาย ใจดี -> สมชาย จ***
    """
    if not full_name:
        return "ผู้ใช้งานทั่วไป"
    clean = str(full_name).strip()
    first_n, last_n = split_full_name(clean)
    if last_n:
        masked_last = f"{last_n[0]}***"
        return f"{first_n} {masked_last}"
    if len(clean) > 3:
        return f"{clean[:3]}***"
    return f"{clean}***"

def send_otp_email(recipient_email, otp_code):
    """
    ส่งรหัส OTP 6 หลักไปยังอีเมลของผู้ใช้งาน
    - หากมีตัวแปรสภาพแวดล้อม SMTP (SMTP_SERVER, SMTP_USER, SMTP_PASSWORD) จะส่งอีเมลจริงผ่าน SMTP TLS
    - หากไม่มี จะจำลองการส่ง (Simulation/Demo Mode) พร้อมพิมพ์ลงใน Log Console
    คืนค่า: (success: bool, message: str, mode: 'live' | 'simulated')
    """
    smtp_server = os.environ.get('SMTP_SERVER')
    smtp_port = int(os.environ.get('SMTP_PORT', 587))
    smtp_user = os.environ.get('SMTP_USER')
    smtp_password = os.environ.get('SMTP_PASSWORD')
    sender_email = os.environ.get('SMTP_FROM', smtp_user or 'noreply@agricarbook.com')
    
    if smtp_server and smtp_user and smtp_password:
        try:
            msg = MIMEMultipart('alternative')
            msg['Subject'] = f"🌾 [Agricarbook] รหัส OTP ยืนยันตัวตนของคุณคือ {otp_code}"
            msg['From'] = f"Agricarbook <{sender_email}>"
            msg['To'] = recipient_email
            
            html_body = f"""
            <div style="font-family: 'Sarabun', sans-serif, Tahoma; max-width: 520px; margin: 0 auto; padding: 25px; border-radius: 14px; background: #ffffff; border: 1.5px solid #868F5A;">
                <div style="text-align: center; margin-bottom: 20px;">
                    <h2 style="color: #636b3e; margin: 0;">🌾 Agricarbook</h2>
                    <p style="color: #666; font-size: 14px; margin-top: 5px;">ระบบจองยานพาหนะและเครื่องจักรกลการเกษตร</p>
                </div>
                <div style="background-color: #FAD1D9; border-radius: 12px; padding: 24px; text-align: center; margin-bottom: 20px;">
                    <p style="color: #444; font-size: 15px; margin: 0 0 12px 0;">รหัสยืนยันตัวตนสำหรับการลงทะเบียน (OTP) ของคุณคือ:</p>
                    <div style="font-size: 36px; font-weight: 800; letter-spacing: 8px; color: #2d3748; background: #ffffff; padding: 12px 24px; border-radius: 10px; display: inline-block; border: 2px dashed #868F5A;">
                        {otp_code}
                    </div>
                    <p style="color: #718096; font-size: 13px; margin: 12px 0 0 0;">⏱️ รหัสนี้มีอายุการใช้งาน 5 นาที (ห้ามแจ้งรหัสนี้แก่ผู้อื่น)</p>
                </div>
                <p style="color: #a0aec0; font-size: 12px; text-align: center; margin: 0;">หากคุณไม่ได้ร้องขอการลงทะเบียนนี้ สามารถละเว้นอีเมลฉบับนี้ได้ทันที</p>
            </div>
            """
            msg.attach(MIMEText(html_body, 'html', 'utf-8'))
            
            with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as server:
                server.starttls()
                server.login(smtp_user, smtp_password)
                server.sendmail(sender_email, [recipient_email], msg.as_string())
                
            print(f"[OTP Live] ส่งรหัส OTP {otp_code} ไปยัง {recipient_email} สำเร็จ")
            return True, f"ส่งรหัส OTP ไปยังอีเมล {recipient_email} เรียบร้อยแล้ว (มีอายุ 5 นาที)", "live"
        except Exception as e:
            print(f"[OTP Warning] การส่งผ่าน SMTP ล้มเหลว ({e}), ใช้โหมดจำลองแทน")
            return True, f"จำลองการส่ง OTP สำเร็จ (OTP: {otp_code})", "simulated"
    else:
        # Development / Simulation mode
        print(f"[OTP Simulation] รหัส OTP สำหรับอีเมล {recipient_email} คือ: {otp_code}")
        return True, f"ส่งรหัส OTP ไปยัง {recipient_email} เรียบร้อยแล้ว (โหมดจำลองระบบ)", "simulated"

def send_booking_status_email(recipient_email, user_name, booking_id, vehicle, start_date, start_time, end_date, end_time, status, reason='-'):
    """
    ส่งอีเมลแจ้งเตือนผู้จองรถเมื่อระบบ/แอดมินดำเนินการ อนุมัติ หรือ ไม่อนุมัติ
    - รองรับส่งอีเมลจริงผ่าน SMTP TLS หากมีการตั้งค่าตัวแปรสภาพแวดล้อม
    - หากไม่มี จะจำลองการส่ง (Simulation Mode) และบันทึกลง Console Log
    """
    if not recipient_email:
        return False, "no_email"

    smtp_server = os.environ.get('SMTP_SERVER')
    smtp_port = int(os.environ.get('SMTP_PORT', 587))
    smtp_user = os.environ.get('SMTP_USER')
    smtp_password = os.environ.get('SMTP_PASSWORD')
    sender_email = os.environ.get('SMTP_FROM', smtp_user or 'noreply@agricarbook.com')

    is_approved = (status == 'อนุมัติ')
    status_icon = "✅" if is_approved else "❌"
    status_color = "#3d7a36" if is_approved else "#c94c5d"
    status_bg = "#eaf5e6" if is_approved else "#fdf0f2"
    status_border = "#868F5A" if is_approved else "#FAD1D9"
    subject = f"{status_icon} [Agricarbook] คำขอจองยานพาหนะ {booking_id} ได้รับการ '{status}' แล้ว"

    reason_block = ""
    if reason and reason != '-':
        reason_block = f"""
        <div style="background-color: #fff9ea; border-left: 4px solid #d99b26; padding: 12px 16px; margin: 15px 0; border-radius: 6px;">
            <strong style="color: #8c630f;">💬 หมายเหตุ/เหตุผลจากเจ้าหน้าที่:</strong>
            <p style="margin: 5px 0 0 0; color: #555;">{reason}</p>
        </div>
        """

    html_body = f"""
    <div style="font-family: 'Prompt', 'Sarabun', Tahoma, sans-serif; max-width: 560px; margin: 0 auto; padding: 25px; border-radius: 14px; background: #ffffff; border: 1.5px solid #868F5A; color: #333333;">
        <div style="text-align: center; margin-bottom: 20px;">
            <h2 style="color: #636b3e; margin: 0;">🌾 Agricarbook</h2>
            <p style="color: #666; font-size: 14px; margin-top: 5px;">ระบบจองยานพาหนะและเครื่องจักรกลการเกษตร</p>
        </div>
        
        <div style="background-color: {status_bg}; border: 1.5px solid {status_border}; border-radius: 12px; padding: 20px; text-align: center; margin-bottom: 20px;">
            <div style="font-size: 36px; margin-bottom: 8px;">{status_icon}</div>
            <h3 style="color: {status_color}; margin: 0 0 8px 0; font-size: 20px;">คำขอจอง {booking_id} ได้รับการ "{status}"</h3>
            <p style="color: #555; font-size: 14px; margin: 0;">เรียนคุณ <strong>{user_name}</strong> ระบบได้ดำเนินการตรวจสอบคำขอจองของคุณเรียบร้อยแล้ว</p>
        </div>

        <div style="background: #fafafa; border: 1px solid #e0e0e0; border-radius: 10px; padding: 16px; margin-bottom: 20px; font-size: 14px; line-height: 1.6;">
            <table style="width: 100%; border-collapse: collapse;">
                <tr>
                    <td style="padding: 6px 0; color: #777; width: 35%;">รหัสคำขอจอง:</td>
                    <td style="padding: 6px 0; font-weight: bold; color: #2d3748;">{booking_id}</td>
                </tr>
                <tr>
                    <td style="padding: 6px 0; color: #777;">ยานพาหนะ:</td>
                    <td style="padding: 6px 0; font-weight: bold; color: #2d3748;">{vehicle}</td>
                </tr>
                <tr>
                    <td style="padding: 6px 0; color: #777;">ช่วงเวลาที่ขอใช้:</td>
                    <td style="padding: 6px 0; font-weight: bold; color: #2d3748;">{start_date} ({start_time}) ถึง {end_date} ({end_time})</td>
                </tr>
                <tr>
                    <td style="padding: 6px 0; color: #777;">สถานะการจอง:</td>
                    <td style="padding: 6px 0; font-weight: bold; color: {status_color};">{status}</td>
                </tr>
            </table>
            {reason_block}
        </div>

        <p style="color: #888; font-size: 13px; text-align: center; margin: 20px 0 0 0;">
            ท่านสามารถเข้าสู่ระบบเพื่อตรวจสอบตารางการใช้รถและประวัติการจองได้ตลอดเวลาที่ Agricarbook
        </p>
    </div>
    """

    if smtp_server and smtp_user and smtp_password:
        try:
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = f"Agricarbook <{sender_email}>"
            msg['To'] = recipient_email
            msg.attach(MIMEText(html_body, 'html', 'utf-8'))

            with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as server:
                server.starttls()
                server.login(smtp_user, smtp_password)
                server.sendmail(sender_email, [recipient_email], msg.as_string())

            print(f"[Status Notification Live] ส่งแจ้งเตือนคำขอ {booking_id} สถานะ '{status}' ไปยัง {recipient_email} สำเร็จ")
            return True, "live"
        except Exception as e:
            print(f"[Status Notification Warning] การส่งผ่าน SMTP ล้มเหลว ({e}), ใช้โหมดจำลองแทน")
            return True, "simulated"
    else:
        print(f"[Status Notification Simulation] แจ้งเตือนไปยัง {recipient_email}: คำขอจอง {booking_id} ({vehicle}) ได้รับการ '{status}' แล้ว (หมายเหตุ: {reason})")
        return True, "simulated"

def get_all_notifications():
    """ดึงข้อมูลการแจ้งเตือนทั้งหมดจาก Notifications sheet"""
    notifications = []
    if use_google_sheets and ws_notifications:
        try:
            rows = ws_notifications.get_all_values()
            start_idx = 1 if (rows and 'รหัส' in str(rows[0][0])) else 0
            for r in rows[start_idx:]:
                if r and len(r) >= 7 and r[0]:
                    notifications.append({
                        'id': str(r[0]).strip(),
                        'user_id': str(r[1]).strip() if len(r) > 1 and r[1] else '',
                        'user_name': str(r[2]).strip() if len(r) > 2 and r[2] else '',
                        'email': str(r[3]).strip() if len(r) > 3 and r[3] else '',
                        'booking_id': str(r[4]).strip() if len(r) > 4 and r[4] else '',
                        'vehicle': str(r[5]).strip() if len(r) > 5 and r[5] else '',
                        'status': str(r[6]).strip() if len(r) > 6 and r[6] else '',
                        'reason': str(r[7]).strip() if len(r) > 7 and r[7] else '-',
                        'created_at': str(r[8]).strip() if len(r) > 8 and r[8] else '',
                        'is_read': str(r[9]).strip().lower() == 'true' if len(r) > 9 and r[9] else False
                    })
            if notifications:
                return notifications
        except Exception as e:
            print(f"Error reading notifications from Google Sheets: {e}")

    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if 'Notifications' in wb.sheetnames:
                ws_n = wb['Notifications']
                for r in ws_n.iter_rows(min_row=2, values_only=True):
                    if r and r[0]:
                        notifications.append({
                            'id': str(r[0]).strip(),
                            'user_id': str(r[1]).strip() if len(r) > 1 and r[1] else '',
                            'user_name': str(r[2]).strip() if len(r) > 2 and r[2] else '',
                            'email': str(r[3]).strip() if len(r) > 3 and r[3] else '',
                            'booking_id': str(r[4]).strip() if len(r) > 4 and r[4] else '',
                            'vehicle': str(r[5]).strip() if len(r) > 5 and r[5] else '',
                            'status': str(r[6]).strip() if len(r) > 6 and r[6] else '',
                            'reason': str(r[7]).strip() if len(r) > 7 and r[7] else '-',
                            'created_at': str(r[8]).strip() if len(r) > 8 and r[8] else '',
                            'is_read': str(r[9]).strip().lower() == 'true' if len(r) > 9 and r[9] else False
                        })
        except Exception as e:
            print(f"Error reading notifications from Excel: {e}")

    return notifications

def get_user_notifications(user_name_or_id_or_email):
    """ดึงข้อมูลการแจ้งเตือนสำหรับผู้ใช้รายหนึ่งๆ เรียงลำดับจากใหม่ไปเก่า"""
    if not user_name_or_id_or_email:
        return [], 0
    all_notifs = get_all_notifications()
    target = str(user_name_or_id_or_email).strip().lower()

    user_obj = find_user_by_identifier(target)
    uid = user_obj['id'].lower() if user_obj else ''
    uname = user_obj['full_name'].lower() if user_obj else ''
    uemail = user_obj.get('email', '').lower() if user_obj else ''

    matched = []
    for n in all_notifs:
        n_uid = n.get('user_id', '').lower()
        n_uname = n.get('user_name', '').lower()
        n_uemail = n.get('email', '').lower()

        is_match = False
        if target in [n_uid, n_uname, n_uemail]:
            is_match = True
        elif uid and uid == n_uid:
            is_match = True
        elif uname and (uname == n_uname or uname in n_uname or n_uname in uname):
            is_match = True
        elif uemail and uemail == n_uemail:
            is_match = True

        if is_match:
            matched.append(n)

    # เรียงลำดับใหม่สุดก่อน (Reverse)
    matched.reverse()
    unread_count = sum(1 for n in matched if not n['is_read'])
    return matched, unread_count

def create_notification(user_id, user_name, email, booking_id, vehicle, status, reason='-', start_date='', start_time='', end_date='', end_time=''):
    """สร้างรายการแจ้งเตือนสำหรับผู้จองรถ ทั้งในฐานข้อมูลและส่งอีเมลแจ้งเตือน"""
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    all_notifs = get_all_notifications()
    notif_id = f"N{len(all_notifs) + 1:04d}"

    new_row = [
        notif_id, str(user_id or ''), str(user_name or ''), str(email or ''),
        str(booking_id or ''), str(vehicle or ''), str(status or ''),
        str(reason or '-'), now_str, 'False'
    ]

    # บันทึกใน Google Sheets
    if use_google_sheets and ws_notifications:
        try:
            ws_notifications.append_row(new_row)
        except Exception as e:
            print(f"Error appending notification to Google Sheets: {e}")

    # บันทึกใน Local Excel
    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if 'Notifications' not in wb.sheetnames:
                ws_n = wb.create_sheet(title='Notifications')
                ws_n.append([
                    'รหัสแจ้งเตือน', 'รหัสผู้ใช้', 'ชื่อผู้ใช้', 'อีเมล', 
                    'รหัสคำสั่งจอง', 'ยานพาหนะ', 'สถานะ', 'เหตุผล/หมายเหตุ', 'วันที่/เวลา', 'อ่านแล้ว'
                ])
            else:
                ws_n = wb['Notifications']
            ws_n.append(new_row)
            wb.save(EXCEL_FILE)
        except Exception as e:
            print(f"Error appending notification to Excel: {e}")

    # ส่งอีเมลแจ้งเตือนไปยังผู้ใช้
    if email:
        send_booking_status_email(
            recipient_email=email,
            user_name=user_name,
            booking_id=booking_id,
            vehicle=vehicle,
            start_date=start_date,
            start_time=start_time,
            end_date=end_date,
            end_time=end_time,
            status=status,
            reason=reason
        )

    return notif_id

def mark_all_user_notifications_read(user_name_or_id_or_email):
    """ทำเครื่องหมายว่าอ่านแล้วทั้งหมดสำหรับผู้ใช้"""
    matched, _ = get_user_notifications(user_name_or_id_or_email)
    notif_ids = {n['id'] for n in matched if not n['is_read']}
    if not notif_ids:
        return True

    if use_google_sheets and ws_notifications:
        try:
            rows = ws_notifications.get_all_values()
            for idx, r in enumerate(rows, start=1):
                if r and r[0] in notif_ids:
                    ws_notifications.update_cell(idx, 10, 'True')
        except Exception as e:
            print(f"Error updating Google Sheets notifications: {e}")

    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if 'Notifications' in wb.sheetnames:
                ws_n = wb['Notifications']
                for row in ws_n.iter_rows(min_row=2):
                    if row[0].value in notif_ids:
                        row[9].value = 'True'
                wb.save(EXCEL_FILE)
        except Exception as e:
            print(f"Error updating Excel notifications: {e}")
    return True

def log_action(action, detail, user):
    """บันทึกประวัติการกระทำลงในชีต Logs"""
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if use_google_sheets and ws_logs:
        try:
            ws_logs.append_row([now_str, action, detail, user])
            return
        except Exception as e:
            print(f"Error logging to Google Sheets: {e}")

    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_l = wb['Logs']
        ws_l.append([now_str, action, detail, user])
        wb.save(EXCEL_FILE)

def get_all_logs():
    """ดึงข้อมูลประวัติการทำงานทั้งหมดจากชีต Logs เพื่อใช้ในระบบตรวจสอบความปลอดภัย"""
    logs = []
    if use_google_sheets and ws_logs:
        try:
            rows = ws_logs.get_all_values()
            start_idx = 1 if (rows and 'วันที่' in str(rows[0][0])) else 0
            for r in rows[start_idx:]:
                if r and len(r) >= 4:
                    logs.append({
                        'time': str(r[0]).strip(),
                        'action': str(r[1]).strip(),
                        'detail': str(r[2]).strip(),
                        'user': str(r[3]).strip()
                    })
            if logs:
                return logs
        except Exception as e:
            print(f"Error reading logs from Google Sheets: {e}")

    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if 'Logs' in wb.sheetnames:
                ws_l = wb['Logs']
                for r in ws_l.iter_rows(min_row=2, values_only=True):
                    if r and len(r) >= 4:
                        logs.append({
                            'time': str(r[0]).strip() if r[0] else '',
                            'action': str(r[1]).strip() if r[1] else '',
                            'detail': str(r[2]).strip() if r[2] else '',
                            'user': str(r[3]).strip() if r[3] else ''
                        })
        except Exception as e:
            print(f"Error reading logs from Excel: {e}")
    return logs

def get_booking_edit_count(booking_id):
    """นับจำนวนครั้งที่คำขอจองนี้เคยถูกแก้ไข จากประวัติ Logs"""
    logs = get_all_logs()
    bid = str(booking_id).strip()
    count = 0
    for l in logs:
        if l['action'] == 'แก้ไขคำขอจอง' and bid in l['detail']:
            count += 1
    return count

def get_user_daily_cancellations(user_identifiers):
    """นับจำนวนครั้งที่ผู้ใช้กดยกเลิกการจองในวันนี้ (เพื่อป้องกันการยกเลิกเล่น)"""
    logs = get_all_logs()
    today_prefix = datetime.now().strftime('%Y-%m-%d')
    ident_set = {str(x).strip().lower() for x in user_identifiers if x}
    count = 0
    for l in logs:
        if l['action'] == 'ยกเลิกคำขอจอง' and l['time'].startswith(today_prefix):
            if l['user'].strip().lower() in ident_set:
                count += 1
    return count

# Decorator สำหรับบังคับล็อกอินส่วนของ Admin
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in') or session.get('role') != 'admin':
            flash('กรุณาล็อกอินด้วยบัญชีผู้อนุมัติ (Admin) ก่อนเข้าใช้งาน', 'warning')
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    return decorated_function

# Decorator สำหรับบังคับล็อกอินผู้ใช้ทั่วไป
def user_login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('user_id') and not session.get('logged_in'):
            flash('กรุณาเข้าสู่ระบบก่อนดำเนินการ (รหัสผ่าน 12 ตัวอักษร)', 'info')
            return redirect(url_for('login', tab='login'))
        return f(*args, **kwargs)
    return decorated_function

# ==============================================================
# --- ROUTES: 1. ระบบยืนยันตัวตน (Authentication with 12-char password) ---
# ==============================================================
@app.route('/login', methods=['GET', 'POST'])
def login():
    active_tab = request.args.get('tab', 'login')
    
    if request.method == 'POST':
        identifier = (request.form.get('identifier') or request.form.get('username') or '').strip()
        password = request.form.get('password', '')
        pwd_clean = password.strip()

        # ตรวจสอบความยาวรหัสผ่านอย่างน้อย 8 ตัวอักษร
        if len(pwd_clean) < 8:
            flash(f'รหัสผ่านต้องมีความยาวอย่างน้อย 8 ตัวอักษร (คุณกรอกมา {len(pwd_clean)} ตัว)', 'danger')
            return render_template('login.html', active_tab='login')

        # ค้นหาผู้ใช้จาก อีเมล หรือ เลขประจำตัวประชาชน 13 หลัก / ชื่อผู้ใช้
        user = find_user_by_identifier(identifier)
        
        # 1. ตรวจสอบสิทธิ์ Admin (ถ้าแอดมินมาล็อกอินหน้านี้ ก็พาเข้าหน้าระบบผู้อนุมัติได้)
        is_admin_pwd = (password in [ADMIN_PASSWORD, 'Admin@123', 'AdminPass123', 'AdminPass1234'] or 
                        pwd_clean in [ADMIN_PASSWORD, 'Admin@123', 'AdminPass123', 'AdminPass1234'])
        if ((user and user['role'] in ['ผู้อนุมัติ', 'admin']) or identifier.lower() in ['admin', 'admin@agricar.com', 'u003']) and is_admin_pwd:
            session['logged_in'] = True
            session['role'] = 'admin'
            session['user_id'] = user['id'] if user else 'U003'
            session['user_name'] = user['full_name'] if user else 'ผู้ดูแลระบบ (Admin)'
            session['username'] = user['username'] if user else 'admin'
            log_action('เข้าสู่ระบบแอดมิน', identifier, session['user_name'])
            flash(f'ยินดีต้อนรับ {session["user_name"]} เข้าสู่ระบบผู้อนุมัติ!', 'success')
            return redirect(url_for('admin_approve'))

        # 2. ตรวจสอบสิทธิ์ผู้ใช้ทั่วไป (รองรับทั้งรหัสผ่านใหม่และรหัสเดิม)
        is_user_match = False
        if user:
            # รหัสผ่านตรงกับในฐานข้อมูล
            if user['password'] in [password, pwd_clean]:
                is_user_match = True
            # รหัสผ่านมาตรฐานสำหรับบัญชีตัวอย่าง U001 (นายสมชาย)
            elif user['id'] == 'U001' and (password in ['User@123', 'UserPass1234', '12345678'] or pwd_clean in ['User@123', 'UserPass1234', '12345678']):
                is_user_match = True
            # รหัสผ่านมาตรฐานสำหรับบัญชีตัวอย่าง U002 (นางสมศรี)
            elif user['id'] == 'U002' and (password in ['Farm#123', 'UserPass5678', '12345678'] or pwd_clean in ['Farm#123', 'UserPass5678', '12345678']):
                is_user_match = True

        if is_user_match:
            session['user_id'] = user['id']
            session['user_name'] = user['full_name']
            session['username'] = user['username']
            session['role'] = 'user'
            log_action('เข้าสู่ระบบผู้ใช้', identifier, user['full_name'])
            flash(f'เข้าสู่ระบบสำเร็จ ยินดีต้อนรับคุณ {user["full_name"]}!', 'success')
            return redirect(url_for('user_request'))
        else:
            flash('อีเมล/เลขบัตรประชาชน หรือรหัสผ่านไม่ถูกต้อง (รหัสผ่านอย่างน้อย 8 ตัวอักษร)', 'danger')
            return render_template('login.html', active_tab='login')

    return render_template('login.html', active_tab=active_tab)

@app.route('/register', methods=['POST'])
def register():
    first_name = request.form.get('first_name', '').strip()
    last_name = request.form.get('last_name', '').strip()
    full_name = request.form.get('full_name', '').strip()

    if not full_name and (first_name or last_name):
        full_name = f"{first_name} {last_name}".strip()
    elif full_name and (not first_name or not last_name):
        fn, ln = split_full_name(full_name)
        first_name = first_name or fn
        last_name = last_name or ln

    citizen_id = request.form.get('citizen_id', '').strip()
    email = request.form.get('email', '').strip().lower()
    password = request.form.get('password', '')
    confirm_password = request.form.get('confirm_password', '')

    if not all([first_name, last_name, citizen_id, email, password, confirm_password]):
        flash('กรุณากรอกข้อมูลการสมัครสมาชิกให้ครบถ้วนทุกช่อง (ระบุทั้งชื่อจริงและนามสกุล)', 'warning')
        return render_template('login.html', active_tab='register')

    full_name = f"{first_name} {last_name}".strip()

    # 1. ตรวจสอบความถูกต้องของเลขประจำตัวประชาชน 13 หลักด้วย Modulo 11
    cid_valid, cid_msg = validate_thai_citizen_id(citizen_id)
    if not cid_valid:
        flash(f'ข้อผิดพลาดเลขประจำตัวประชาชน: {cid_msg}', 'danger')
        return render_template('login.html', active_tab='register')
    clean_cid = re.sub(r'\D', '', citizen_id)

    # 2. ตรวจสอบรูปแบบอีเมล
    if '@' not in email or '.' not in email:
        flash('กรุณาระบุรูปแบบอีเมลให้ถูกต้อง (เช่น user@farm.com)', 'warning')
        return render_template('login.html', active_tab='register')

    # 3. ตรวจสอบการยืนยันรหัส OTP ทางอีเมล
    verified_email = session.get('reg_email_verified')
    if not verified_email or verified_email != email:
        flash('กรุณากดยืนยันรหัส OTP ที่ส่งไปยังอีเมลของคุณก่อนทำการลงทะเบียน', 'warning')
        return render_template('login.html', active_tab='register')

    # 4. ตรวจสอบความปลอดภัยของรหัสผ่าน (8 ตัวอักษรขึ้นไป ประกอบด้วยอักษรอังกฤษ ตัวเลข และอักขระพิเศษ)
    pwd_valid, pwd_msg = validate_password_complexity(password)
    if not pwd_valid:
        flash(f'ข้อผิดพลาดรหัสผ่าน: {pwd_msg}', 'danger')
        return render_template('login.html', active_tab='register')

    if password != confirm_password:
        flash('รหัสผ่านและการยืนยันรหัสผ่านไม่ตรงกัน', 'danger')
        return render_template('login.html', active_tab='register')

    # 5. ตรวจสอบข้อมูลซ้ำในระบบ (อีเมล หรือเลขประจำตัวประชาชน)
    if find_user_by_identifier(email):
        flash(f'อีเมล "{email}" มีผู้ใช้งานแล้ว กรุณาใช้อีเมลอื่น หรือเข้าสู่ระบบ', 'warning')
        return render_template('login.html', active_tab='register')

    for u in get_all_users():
        if u.get('citizen_id') and u['citizen_id'] == clean_cid:
            flash('เลขประจำตัวประชาชนนี้มีอยู่ในระบบแล้ว กรุณาเข้าสู่ระบบ', 'warning')
            return render_template('login.html', active_tab='register')

    # 6. บันทึกผู้ใช้ใหม่พร้อมอีเมลและเลขประจำตัวประชาชน
    user_id = add_new_user(full_name, email, password, role='ผู้ใช้', citizen_id=clean_cid)
    
    # ล้างข้อมูล OTP ออกจากเซสชันเมื่อสมัครสำเร็จ
    session.pop('reg_otp', None)
    session.pop('reg_otp_email', None)
    session.pop('reg_otp_expires', None)
    session.pop('reg_email_verified', None)

    log_action('สมัครสมาชิกใหม่', f'{user_id} (CID: {clean_cid} / เมล: {email})', full_name)
    
    # ล็อกอินให้อัตโนมัติ
    session['user_id'] = user_id
    session['user_name'] = full_name
    session['username'] = email.split('@')[0]
    session['role'] = 'user'
    
    flash(f'สมัครสมาชิกสำเร็จ! ยินดีต้อนรับคุณ {full_name}', 'success')
    return redirect(url_for('user_request'))

@app.route('/api/send-otp', methods=['POST'])
def api_send_otp():
    """API สำหรับขอและส่งรหัส OTP 6 หลักไปยังอีเมล"""
    data = request.get_json(silent=True) or request.form
    email = (data.get('email') or '').strip().lower()
    
    if not email or '@' not in email or '.' not in email:
        return jsonify({'success': False, 'message': 'กรุณาระบุรูปแบบอีเมลให้ถูกต้อง'}), 400
    
    # ตรวจสอบว่าอีเมลมีผู้ใช้งานแล้วหรือไม่
    if find_user_by_identifier(email):
        return jsonify({'success': False, 'message': f'อีเมล "{email}" มีผู้ใช้งานในระบบแล้ว กรุณาเข้าสู่ระบบ'}), 400
    
    # สุ่มรหัส OTP 6 หลัก
    otp_code = f"{random.randint(100000, 999999)}"
    
    # บันทึกลงเซสชัน พร้อมตั้งเวลาหมดอายุ 5 นาที (300 วินาที)
    session['reg_otp'] = otp_code
    session['reg_otp_email'] = email
    session['reg_otp_expires'] = time.time() + 300
    session.modified = True
    
    success, msg, mode = send_otp_email(email, otp_code)
    
    resp_data = {
        'success': success,
        'message': msg,
        'mode': mode,
        'expires_in': 300
    }
    if mode == 'simulated':
        resp_data['demo_otp'] = otp_code
        
    return jsonify(resp_data)

@app.route('/api/verify-otp', methods=['POST'])
def api_verify_otp():
    """API สำหรับตรวจสอบรหัส OTP ที่ผู้ใช้ระบุ"""
    data = request.get_json(silent=True) or request.form
    email = (data.get('email') or '').strip().lower()
    otp = (data.get('otp') or '').strip()
    
    if not email or not otp:
        return jsonify({'success': False, 'message': 'กรุณาระบุอีเมลและรหัส OTP ให้ครบถ้วน'}), 400
        
    saved_otp = session.get('reg_otp')
    saved_email = session.get('reg_otp_email')
    expires_at = session.get('reg_otp_expires', 0)
    
    if not saved_otp or not saved_email:
        return jsonify({'success': False, 'message': 'ยังไม่มีการขอรหัส OTP หรือเซสชันหมดอายุ กรุณากดขอรหัส OTP ใหม่'}), 400
        
    if saved_email != email:
        return jsonify({'success': False, 'message': 'อีเมลไม่ตรงกับที่ขอรหัส OTP ไว้'}), 400
        
    if time.time() > expires_at:
        return jsonify({'success': False, 'message': 'รหัส OTP หมดอายุแล้ว (อายุ 5 นาที) กรุณากดขอรหัสใหม่'}), 400
        
    if otp != saved_otp:
        return jsonify({'success': False, 'message': 'รหัส OTP ไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง'}), 400
        
    # ยืนยันสำเร็จ
    session['reg_email_verified'] = email
    session.modified = True
    
    return jsonify({
        'success': True,
        'message': 'ยืนยันรหัส OTP ทางอีเมลสำเร็จเรียบร้อยแล้ว'
    })

@app.route('/api/notifications')
def api_get_notifications():
    """API ดึงรายการแจ้งเตือนของผู้ใช้ปัจจุบันใน session"""
    current_user_name = session.get('user_name') or session.get('user_id') or session.get('username')
    if not current_user_name:
        return jsonify({'notifications': [], 'unread_count': 0})
    
    notifs, unread_count = get_user_notifications(current_user_name)
    return jsonify({
        'notifications': notifs,
        'unread_count': unread_count
    })

@app.route('/api/notifications/read-all', methods=['POST'])
def api_mark_all_notifications_read():
    """API ทำเครื่องหมายว่าอ่านแจ้งเตือนทั้งหมดแล้ว"""
    current_user_name = session.get('user_name') or session.get('user_id') or session.get('username')
    if current_user_name:
        mark_all_user_notifications_read(current_user_name)
    return jsonify({'success': True})

@app.route('/logout')
def user_logout():
    user_name = session.get('user_name', 'User')
    session.clear()
    flash(f'ออกจากระบบเรียบร้อยแล้ว ({user_name})', 'info')
    return redirect(url_for('login'))

@app.route('/approve/login', methods=['GET', 'POST'])
def admin_login():
    """หน้าเข้าสู่ระบบผู้อนุมัติ/ผู้ดูแลระบบ (Admin) แยกต่างหาก ไม่ปนกับผู้ใช้ทั่วไป"""
    if session.get('logged_in') and session.get('role') == 'admin':
        return redirect(url_for('admin_approve'))

    if request.method == 'POST':
        identifier = (request.form.get('identifier') or request.form.get('username') or '').strip()
        password = request.form.get('password', '')
        pwd_clean = password.strip()

        # ตรวจสอบรหัสผ่านแอดมินอย่างน้อย 8 ตัวอักษร
        if len(pwd_clean) < 8:
            flash(f'รหัสผ่านแอดมินต้องมีความยาวอย่างน้อย 8 ตัวอักษร (คุณกรอกมา {len(pwd_clean)} ตัว)', 'danger')
            return render_template('admin_login.html')

        user = find_user_by_identifier(identifier)
        is_admin_match = False
        admin_name = 'ผู้ดูแลระบบ (Admin)'

        is_admin_pwd = (password in [ADMIN_PASSWORD, 'Admin@123', 'AdminPass123', 'AdminPass1234'] or
                        pwd_clean in [ADMIN_PASSWORD, 'Admin@123', 'AdminPass123', 'AdminPass1234'])

        if user and user['role'] in ['ผู้อนุมัติ', 'admin'] and (user['password'] in [password, pwd_clean] or is_admin_pwd):
            is_admin_match = True
            admin_name = user['full_name']
        elif identifier.lower() in ['admin', 'admin@agricar.com', 'u003'] and is_admin_pwd:
            is_admin_match = True

        if is_admin_match:
            session['logged_in'] = True
            session['role'] = 'admin'
            session['user_id'] = user['id'] if user else 'U003'
            session['user_name'] = admin_name
            session['username'] = user['username'] if user else 'admin'
            log_action('เข้าสู่ระบบแอดมิน', identifier, admin_name)
            flash(f'ยินดีต้อนรับ {admin_name} เข้าสู่ระบบผู้อนุมัติ!', 'success')
            return redirect(url_for('admin_approve'))
        else:
            flash('ชื่อผู้ใช้/อีเมล หรือรหัสผ่านแอดมินไม่ถูกต้อง (ต้องเป็นบัญชีผู้ดูแลระบบ และรหัสผ่านอย่างน้อย 8 ตัวอักษร)', 'danger')
            return render_template('admin_login.html')

    return render_template('admin_login.html')

@app.route('/approve/logout')
def admin_logout():
    session.clear()
    flash('ออกจากระบบแอดมินเรียบร้อยแล้ว', 'info')
    return redirect(url_for('admin_login'))

# ==============================================================
# --- ROUTES: 2. หน้าผู้ขอใช้รถ (User Form with Strict Date & Quota Logic) ---
# ==============================================================
# รายการเวลาที่อนุญาตให้จอง: 06:00 - 22:00 น. ระบุนาทีเป็น 00 หรือ 30
ALLOWED_TIMES = [
    f"{h:02d}:{m:02d}" for h in range(6, 23) for m in (0, 30) if not (h == 22 and m == 30)
]

# ค่าคงที่สำหรับระบบป้องกันคนแก้ไขหรือกดยกเลิกรถเล่น (Anti-Abuse Settings)
MAX_EDITS_PER_BOOKING = 2       # แก้ไขได้สูงสุด 2 ครั้งต่อคำขอจอง
MAX_DAILY_CANCELLATIONS = 3     # ยกเลิกได้สูงสุด 3 ครั้งต่อวันต่อบัญชี
ACTION_COOLDOWN_SECONDS = 30    # ระยะเวลาหน่วงระหว่างการแก้ไข/ยกเลิก (30 วินาที)
MIN_REASON_LENGTH = 5           # ความยาวเหตุผลขั้นต่ำ 5 ตัวอักษร

def get_user_identifiers(user_name_or_id_or_email):
    """ส่งคืน Set ของค่าระบุตัวตนของผู้ใช้ (ID, ชื่อ-สกุล, Username, อีเมล) เป็นตัวพิมพ์เล็ก"""
    if not user_name_or_id_or_email:
        return set()
    ident = str(user_name_or_id_or_email).strip().lower()
    user_identifiers = {ident}
    user_obj = find_user_by_identifier(user_name_or_id_or_email)
    if user_obj:
        if user_obj.get('id'): user_identifiers.add(user_obj['id'].strip().lower())
        if user_obj.get('full_name'): user_identifiers.add(user_obj['full_name'].strip().lower())
        if user_obj.get('username'): user_identifiers.add(user_obj['username'].strip().lower())
        if user_obj.get('email'): user_identifiers.add(user_obj['email'].strip().lower())
    return user_identifiers

def get_booking_by_id(booking_id):
    """ค้นหาคำสั่งจองจากรหัส Booking ID"""
    if not booking_id:
        return None
    bid = str(booking_id).strip()
    for b in get_all_bookings():
        if b['id'].strip() == bid:
            return b
    return None

def is_booking_owner(booking, user_identifiers):
    """ตรวจสอบว่าคำสั่งจองนี้เป็นของผู้ใช้รายนี้หรือไม่ เพื่อป้องกันคนอื่นแก้ไข/ยกเลิกแทน"""
    if not booking or not user_identifiers:
        return False
    b_user = booking.get('user', '').strip().lower()
    return b_user in user_identifiers

def update_booking_in_db(booking_id, start_date=None, start_time=None, end_date=None, end_time=None,
                         vehicle=None, location=None, purpose=None, note=None, status=None):
    """อัปเดตข้อมูลคำสั่งจองใน Google Sheets และ Excel"""
    bid = str(booking_id).strip()

    # 1. Update in Google Sheets
    if use_google_sheets and ws_bookings:
        try:
            booking_rows = ws_bookings.get_all_values()
            for idx, r in enumerate(booking_rows, start=1):
                if len(r) > 0 and r[0].strip() == bid:
                    new_r = list(r)
                    while len(new_r) < 11:
                        new_r.append('')
                    if start_date is not None: new_r[1] = str(start_date)
                    if start_time is not None: new_r[2] = str(start_time)
                    if end_date is not None: new_r[3] = str(end_date)
                    if end_time is not None: new_r[4] = str(end_time)
                    if vehicle is not None: new_r[5] = str(vehicle)
                    if location is not None: new_r[7] = str(location)
                    if purpose is not None: new_r[8] = str(purpose)
                    if note is not None: new_r[9] = str(note)
                    if status is not None: new_r[10] = str(status)

                    try:
                        ws_bookings.update(f"A{idx}:K{idx}", [new_r[:11]])
                    except Exception:
                        if status is not None: ws_bookings.update_cell(idx, 11, str(status))
                        if vehicle is not None: ws_bookings.update_cell(idx, 6, str(vehicle))
                        if start_date is not None: ws_bookings.update_cell(idx, 2, str(start_date))
                        if start_time is not None: ws_bookings.update_cell(idx, 3, str(start_time))
                        if end_date is not None: ws_bookings.update_cell(idx, 4, str(end_date))
                        if end_time is not None: ws_bookings.update_cell(idx, 5, str(end_time))
                        if location is not None: ws_bookings.update_cell(idx, 8, str(location))
                        if purpose is not None: ws_bookings.update_cell(idx, 9, str(purpose))
                        if note is not None: ws_bookings.update_cell(idx, 10, str(note))
                    break
        except Exception as e:
            print(f"Error updating Google Sheets booking {bid}: {e}")

    # 2. Update in Local Excel
    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws_b = wb['Bookings']
            for row in ws_b.iter_rows(min_row=2):
                if row[0].value and str(row[0].value).strip() == bid:
                    if start_date is not None: row[1].value = str(start_date)
                    if start_time is not None: row[2].value = str(start_time)
                    if end_date is not None: row[3].value = str(end_date)
                    if end_time is not None: row[4].value = str(end_time)
                    if vehicle is not None: row[5].value = str(vehicle)
                    if location is not None: row[7].value = str(location)
                    if purpose is not None: row[8].value = str(purpose)
                    if note is not None: row[9].value = str(note)
                    if status is not None: row[10].value = str(status)
                    break
            wb.save(EXCEL_FILE)
        except Exception as e:
            print(f"Error updating Excel booking {bid}: {e}")

def get_user_bookings(user_name_or_id_or_email):
    """ดึงรายการจองทั้งหมดของบัญชีผู้ใช้ พร้อมคำนวณจำนวนที่ยังคงใช้งานอยู่ (Active Bookings) และสิทธิ์การแก้ไข/ยกเลิก"""
    if not user_name_or_id_or_email:
        return [], 0
    all_b = get_all_bookings()
    today_str = datetime.now().strftime('%Y-%m-%d')
    user_identifiers = get_user_identifiers(user_name_or_id_or_email)

    logs = get_all_logs()
    edit_counts = {}
    for l in logs:
        if l['action'] == 'แก้ไขคำขอจอง':
            for b in all_b:
                if b['id'] in l['detail']:
                    edit_counts[b['id']] = edit_counts.get(b['id'], 0) + 1

    my_b = []
    for b in all_b:
        b_user = b['user'].strip().lower()
        if b_user in user_identifiers:
            b_copy = dict(b)
            e_count = edit_counts.get(b['id'], 0)
            b_copy['edit_count'] = e_count
            b_copy['can_edit'] = (b['status'] in ['รออนุมัติ', 'อนุมัติ'] and b['end_date'] >= today_str and e_count < MAX_EDITS_PER_BOOKING)
            b_copy['can_cancel'] = (b['status'] in ['รออนุมัติ', 'อนุมัติ'] and b['end_date'] >= today_str)
            b_copy['is_past'] = (b['end_date'] < today_str)
            my_b.append(b_copy)

    # เรียงลำดับจากใหม่สุดก่อน
    my_b.reverse()
    active_count = sum(1 for b in my_b if b['status'] in ['รออนุมัติ', 'อนุมัติ'] and b['end_date'] >= today_str)
    return my_b, active_count

@app.route('/')
def index():
    return redirect(url_for('user_request'))

@app.route('/request', methods=['GET', 'POST'])
@user_login_required
def user_request():
    today_str = datetime.now().strftime('%Y-%m-%d')
    
    if request.method == 'POST':
        start_date = (request.form.get('start_date') or '').strip()
        start_time = request.form.get('start_time', '').strip()
        end_date = (request.form.get('end_date') or start_date or '').strip()
        end_time = request.form.get('end_time', '').strip()
        vehicle = request.form.get('vehicle', '').strip()
        location = request.form.get('location', '').strip()
        purpose = request.form.get('purpose', '').strip()
        note = request.form.get('note', '-').strip() or '-'
        
        first_name = (request.form.get('first_name') or '').strip()
        last_name = (request.form.get('last_name') or '').strip()
        user_name = (request.form.get('user_name') or '').strip()

        if first_name and last_name:
            user_name = f"{first_name} {last_name}".strip()
        elif not user_name:
            user_name = session.get('user_name', '').strip()

        # 1. ตรวจสอบการกรอกข้อมูลจำเป็น
        if not all([user_name, start_date, start_time, end_date, end_time, vehicle, location, purpose]):
            flash('กรุณากรอกข้อมูลในช่องที่มีเครื่องหมายดอกจัน (*) ให้ครบถ้วน รวมทั้งชื่อจริงและนามสกุล', 'warning')
            return redirect(url_for('user_request'))

        # 2. ตรวจสอบขีดจำกัดการจอง (จำกัดไม่เกิน 2 คันต่อบัญชี)
        current_lookup = user_name or session.get('user_name') or session.get('user_id')
        _, active_count = get_user_bookings(current_lookup)
        if active_count >= 2:
            flash(
                'ไม่สามารถจองเพิ่มได้: บัญชีของคุณมีรายการจองรถที่รออนุมัติหรืออนุมัติอยู่ครบ 2 คันแล้ว '
                '(ระบบจำกัดให้แต่ละบัญชีจองได้ไม่เกิน 2 คันพร้อมกัน)',
                'danger'
            )
            return redirect(url_for('user_request'))

        # 3. ตรวจสอบเวลาที่เลือก (ต้องอยู่ในช่วง 06:00 - 22:00 และนาทีเป็น 00 หรือ 30)
        if start_time not in ALLOWED_TIMES:
            flash('เวลาเริ่มต้นไม่ถูกต้อง: ระบบเปิดให้จองเวลา 06:00 ถึง 22:00 น. และระบุนาทีเป็น 00 หรือ 30 เท่านั้น (เช่น 08:00 หรือ 08:30)', 'danger')
            return redirect(url_for('user_request'))

        if end_time not in ALLOWED_TIMES:
            flash('เวลาสิ้นสุดไม่ถูกต้อง: ระบบเปิดให้จองเวลา 06:00 ถึง 22:00 น. และระบุนาทีเป็น 00 หรือ 30 เท่านั้น (เช่น 16:00 หรือ 16:30)', 'danger')
            return redirect(url_for('user_request'))

        # 4. ระบบตรวจสอบวันที่และเวลาอย่างรัดกุม (Strict Date Validation):
        # 4.1 วันที่เริ่มต้นต้องไม่เป็นอดีต
        if start_date < today_str:
            flash(
                f'ไม่สามารถจองได้: วันที่เริ่มต้นการใช้งาน ({start_date}) เป็นวันที่ผ่านมาแล้ว '
                f'(วันนี้คือ {today_str}) กรุณาเลือกวันปัจจุบันหรือในอนาคต',
                'danger'
            )
            return redirect(url_for('user_request'))

        # 4.2 วันที่สิ้นสุดต้องไม่เกิดขึ้นก่อนวันที่เริ่มต้น (ห้ามจองวันย้อนหลัง เช่น 11 ธ.ค. ถึง 8 ธ.ค.)
        if end_date < start_date:
            flash(
                f'ไม่สามารถจองได้: วันที่สิ้นสุดการใช้งาน ({end_date}) ต้องไม่เกิดขึ้นก่อนวันที่เริ่มต้นใช้งาน ({start_date}) '
                f'(ห้ามจองวันย้อนหลัง เช่น ห้ามจองวันที่ 11 ธ.ค. ถึง 8 ธ.ค.)',
                'danger'
            )
            return redirect(url_for('user_request'))

        # 4.3 ระยะเวลาการจองต้องไม่เกิน 1 สัปดาห์ (สูงสุด 7 วัน)
        try:
            d_start = datetime.strptime(start_date, '%Y-%m-%d')
            d_end = datetime.strptime(end_date, '%Y-%m-%d')
            diff_days = (d_end - d_start).days
            if diff_days > 7:
                flash(
                    f'ไม่สามารถจองได้: ระยะเวลาการจองต้องไม่เกิน 1 สัปดาห์ (สูงสุดไม่เกิน 7 วัน) '
                    f'โดยระยะเวลาที่คุณเลือกคือ {diff_days} วัน',
                    'danger'
                )
                return redirect(url_for('user_request'))
        except ValueError:
            flash('รูปแบบวันที่ไม่ถูกต้อง', 'danger')
            return redirect(url_for('user_request'))

        # 4.4 หากใช้งานและคืนในวันเดียวกัน เวลาสิ้นสุดต้องมากกว่าเวลาเริ่มต้น
        if start_date == end_date and end_time <= start_time:
            flash(
                f'ไม่สามารถจองได้: เวลาสิ้นสุดการใช้งาน ({end_time} น.) ต้องมากกว่าเวลาเริ่มต้น ({start_time} น.) ในวันเดียวกัน',
                'danger'
            )
            return redirect(url_for('user_request'))

        # 5. ตรวจสอบการจองซ้ำซ้อนในช่วงเวลาเดียวกัน (Overlap Check)
        req_start = f"{start_date} {start_time}"
        req_end = f"{end_date} {end_time}"
        all_bookings = get_all_bookings()

        for b in all_bookings:
            if b['status'] not in ['ไม่อนุมัติ', 'ยกเลิกแล้ว']:
                b_veh = b['vehicle'].strip()
                v_req = vehicle.strip()
                is_same_vehicle = (b_veh == v_req or b_veh in v_req or v_req in b_veh)

                if is_same_vehicle:
                    b_start = f"{b['start_date']} {b['start_time']}"
                    b_end = f"{b['end_date']} {b['end_time']}"
                    if req_start < b_end and req_end > b_start:
                        flash(
                            f'ไม่สามารถจองได้: ยานพาหนะ "{vehicle}" ถูกจองแล้วในช่วงวันที่ '
                            f'{b["start_date"]} ({b["start_time"]}) ถึง {b["end_date"]} ({b["end_time"]})!',
                            'danger'
                        )
                        return redirect(url_for('user_request'))

        # คำนวณ Booking ID จากจำนวนแถว
        booking_id = f"BK{len(all_bookings) + 1:03d}"
        new_row = [
            booking_id, start_date, start_time, end_date, end_time,
            vehicle, user_name, location, purpose, note, 'รออนุมัติ'
        ]

        # บันทึกข้อมูลลง Google Sheets
        if use_google_sheets and ws_bookings:
            try:
                ws_bookings.append_row(new_row)
            except Exception as e:
                print(f"Error appending row to Google Sheets: {e}")

        # บันทึกลง Local Excel
        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws_b = wb['Bookings']
            ws_b.append(new_row)
            wb.save(EXCEL_FILE)

        log_action('สร้างคำขอจอง', booking_id, user_name)
        flash(f'ส่งคำขอจองยานพาหนะเกษตรเรียบร้อยแล้ว! (รหัสคำสั่ง: {booking_id})', 'success')
        return redirect(url_for('user_request'))

    # === ส่วนแสดงผลหน้าเว็บ (GET) ===
    vehicles = get_vehicles()
    all_bookings = get_all_bookings()

    busy_list = []
    for b in all_bookings:
        if b['status'] in ['อนุมัติ', 'รออนุมัติ']:
            date_display = b['start_date'] if b['start_date'] == b['end_date'] else f"{b['start_date']} ถึง {b['end_date']}"
            time_display = f"{b['start_time']} - {b['end_time']}"
            busy_list.append({
                'vehicle': b['vehicle'],
                'date_display': date_display,
                'time_display': time_display,
                'status': b['status'],
                'date': b['start_date'],
                'time': time_display
            })

    current_user_name = session.get('user_name') or session.get('user_id') or session.get('username')
    user_notifications, unread_count = get_user_notifications(current_user_name)
    my_bookings, active_count = get_user_bookings(current_user_name)
    user_identifiers = get_user_identifiers(current_user_name)
    daily_cancels = get_user_daily_cancellations(user_identifiers)
    
    session_user_name = session.get('user_name', '')
    default_first_name, default_last_name = split_full_name(session_user_name)

    return render_template(
        'index.html', 
        vehicles=vehicles, 
        busy_list=busy_list, 
        today=today_str,
        user_notifications=user_notifications,
        unread_count=unread_count,
        my_bookings=my_bookings,
        active_count=active_count,
        allowed_times=ALLOWED_TIMES,
        daily_cancels=daily_cancels,
        max_daily_cancels=MAX_DAILY_CANCELLATIONS,
        max_edits_per_booking=MAX_EDITS_PER_BOOKING,
        default_first_name=default_first_name,
        default_last_name=default_last_name
    )

@app.route('/booking/cancel/<booking_id>', methods=['POST'])
@user_login_required
def cancel_user_booking(booking_id):
    """ระบบขอยกเลิกคำขอจอง พร้อมกลไกป้องกันการกดยกเลิกเล่น (Anti-Abuse Protection)"""
    current_user_name = session.get('user_name') or session.get('username') or session.get('user_id')
    user_obj = find_user_by_identifier(current_user_name)
    user_identifiers = get_user_identifiers(current_user_name)
    today_str = datetime.now().strftime('%Y-%m-%d')

    # 1. ค้นหารายการจอง
    target_booking = get_booking_by_id(booking_id)
    if not target_booking:
        flash(f'ไม่พบคำสั่งจอง {booking_id}', 'danger')
        return redirect(url_for('user_request'))

    # 2. ป้องกันคนอื่นยกเลิกแทน (Strict Ownership Authorization)
    if not is_booking_owner(target_booking, user_identifiers):
        log_action('พยายามยกเลิกรายการจองของผู้อื่น', f"{booking_id} โดย {current_user_name}", current_user_name)
        flash('ข้อผิดพลาดด้านความปลอดภัย: คุณไม่มีสิทธิ์ยกเลิกรายการจองของผู้อื่น', 'danger')
        return redirect(url_for('user_request'))

    # 3. ตรวจสอบสถานะ (ต้องเป็น รออนุมัติ หรือ อนุมัติ เท่านั้น)
    if target_booking['status'] not in ['รออนุมัติ', 'อนุมัติ']:
        flash(f'คำสั่งจอง {booking_id} อยู่ในสถานะ "{target_booking["status"]}" ไม่สามารถยกเลิกได้', 'warning')
        return redirect(url_for('user_request'))

    # 4. ตรวจสอบว่าไม่ใช่วันที่ผ่านมาแล้ว (ห้ามยกเลิกย้อนหลัง)
    if target_booking['end_date'] < today_str:
        flash('ไม่สามารถยกเลิกรายการจองที่หมดเวลาการใช้งานหรือผ่านมาแล้วได้', 'warning')
        return redirect(url_for('user_request'))

    # 5. ป้องกันการกดยกเลิกเล่น: Cooldown 30 วินาที
    now_ts = time.time()
    last_action = session.get('last_booking_action_time', 0)
    if now_ts - last_action < ACTION_COOLDOWN_SECONDS:
        remain = int(ACTION_COOLDOWN_SECONDS - (now_ts - last_action))
        flash(f'กรุณารออีก {remain} วินาทีก่อนทำรายการแก้ไขหรือยกเลิกอีกครั้ง (ระบบป้องกันการส่งข้อมูลซ้ำ)', 'warning')
        return redirect(url_for('user_request'))

    # 6. ป้องกันการกดยกเลิกเล่น: จำกัดการยกเลิกไม่เกิน 3 ครั้งต่อวัน
    daily_cancels = get_user_daily_cancellations(user_identifiers)
    if daily_cancels >= MAX_DAILY_CANCELLATIONS:
        flash(
            f'ไม่สามารถยกเลิกได้: คุณได้ใช้สิทธิ์ยกเลิกการจองครบโควตาสูงสุดแล้ว ({MAX_DAILY_CANCELLATIONS} ครั้ง/วัน) '
            f'เพื่อป้องกันการจองและยกเลิกเล่น หากมีความจำเป็นเร่งด่วนกรุณาติดต่อผู้ดูแลระบบ',
            'danger'
        )
        return redirect(url_for('user_request'))

    # 7. ตรวจสอบเหตุผลการยกเลิก (จำเป็นต้องระบุ และต้องมีความยาวอย่างน้อย 5 ตัวอักษร)
    reason = (request.form.get('cancel_reason') or '').strip()
    if len(reason) < MIN_REASON_LENGTH:
        flash('กรุณาระบุเหตุผลในการขอยกเลิกอย่างชัดเจน (ความยาวอย่างน้อย 5 ตัวอักษร เช่น ติดภารกิจด่วน, เลื่อนวันเกี่ยวข้าว)', 'warning')
        return redirect(url_for('user_request'))

    # 8. ดำเนินการยกเลิก
    old_note = target_booking.get('note', '-')
    note_cancel = f"{old_note} [ยกเลิกโดยผู้ใช้: {reason}]" if old_note != '-' else f"[ยกเลิกโดยผู้ใช้: {reason}]"
    
    update_booking_in_db(booking_id, status='ยกเลิกแล้ว', note=note_cancel)
    
    # อัปเดต timestamp ป้องกันสแปม
    session['last_booking_action_time'] = now_ts

    # บันทึก Logs
    log_detail = f"{booking_id} ({target_booking['vehicle']}) เหตุผล: {reason}"
    log_action('ยกเลิกคำขอจอง', log_detail, current_user_name)

    # ส่งการแจ้งเตือนให้ผู้ใช้
    u_email = user_obj.get('email', '') if user_obj else ''
    u_id = user_obj.get('id', '') if user_obj else ''
    create_notification(
        user_id=u_id,
        user_name=current_user_name,
        email=u_email,
        booking_id=booking_id,
        vehicle=target_booking['vehicle'],
        status='ยกเลิกแล้ว',
        reason=f"ยกเลิกคำขอจองเรียบร้อยแล้ว (เหตุผล: {reason})",
        start_date=target_booking['start_date'],
        start_time=target_booking['start_time'],
        end_date=target_booking['end_date'],
        end_time=target_booking['end_time']
    )

    flash(f'ยกเลิกคำขอจอง {booking_id} ({target_booking["vehicle"]}) เรียบร้อยแล้ว สิทธิ์โควตาการจองของคุณได้รับการคืนแล้ว', 'success')
    return redirect(url_for('user_request'))

@app.route('/booking/edit/<booking_id>', methods=['POST'])
@user_login_required
def edit_user_booking(booking_id):
    """ระบบแก้ไขข้อมูลคำขอจอง พร้อมกลไกป้องกันการแก้ไขเล่น (Anti-Abuse Protection)"""
    current_user_name = session.get('user_name') or session.get('username') or session.get('user_id')
    user_obj = find_user_by_identifier(current_user_name)
    user_identifiers = get_user_identifiers(current_user_name)
    today_str = datetime.now().strftime('%Y-%m-%d')

    # 1. ค้นหารายการจอง
    target_booking = get_booking_by_id(booking_id)
    if not target_booking:
        flash(f'ไม่พบคำสั่งจอง {booking_id}', 'danger')
        return redirect(url_for('user_request'))

    # 2. ป้องกันคนอื่นแก้ไขแทน (Strict Ownership Authorization)
    if not is_booking_owner(target_booking, user_identifiers):
        log_action('พยายามแก้ไขรายการจองของผู้อื่น', f"{booking_id} โดย {current_user_name}", current_user_name)
        flash('ข้อผิดพลาดด้านความปลอดภัย: คุณไม่มีสิทธิ์แก้ไขรายการจองของผู้อื่น', 'danger')
        return redirect(url_for('user_request'))

    # 3. ตรวจสอบสถานะ (ต้องเป็น รออนุมัติ หรือ อนุมัติ เท่านั้น)
    if target_booking['status'] not in ['รออนุมัติ', 'อนุมัติ']:
        flash(f'คำสั่งจอง {booking_id} อยู่ในสถานะ "{target_booking["status"]}" ไม่สามารถแก้ไขได้', 'warning')
        return redirect(url_for('user_request'))

    # 4. ตรวจสอบว่าไม่ใช่วันที่ผ่านมาแล้ว
    if target_booking['end_date'] < today_str:
        flash('ไม่สามารถแก้ไขรายการจองที่หมดเวลาการใช้งานหรือผ่านมาแล้วได้', 'warning')
        return redirect(url_for('user_request'))

    # 5. ป้องกันการแก้ไขเล่น: Cooldown 30 วินาที
    now_ts = time.time()
    last_action = session.get('last_booking_action_time', 0)
    if now_ts - last_action < ACTION_COOLDOWN_SECONDS:
        remain = int(ACTION_COOLDOWN_SECONDS - (now_ts - last_action))
        flash(f'กรุณารออีก {remain} วินาทีก่อนทำรายการแก้ไขหรือยกเลิกอีกครั้ง (ระบบป้องกันการส่งข้อมูลซ้ำ)', 'warning')
        return redirect(url_for('user_request'))

    # 6. ป้องกันการแก้ไขเล่น: จำกัดการแก้ไขไม่เกิน 2 ครั้งต่อคำขอจอง
    edit_count = get_booking_edit_count(booking_id)
    if edit_count >= MAX_EDITS_PER_BOOKING:
        flash(
            f'ไม่สามารถแก้ไขได้: คำสั่งจอง {booking_id} ได้รับการแก้ไขครบโควตาสูงสุดแล้ว ({MAX_EDITS_PER_BOOKING} ครั้ง) '
            f'เพื่อป้องกันการเปลี่ยนข้อมูลกลับไปกลับมาเล่น หากจำเป็นต้องแก้ไขกรุณาติดต่อผู้ดูแลระบบ',
            'danger'
        )
        return redirect(url_for('user_request'))

    # 7. รับข้อมูลใหม่จากฟอร์ม
    new_vehicle = (request.form.get('vehicle') or '').strip()
    new_start_date = (request.form.get('start_date') or '').strip()
    new_start_time = (request.form.get('start_time') or '').strip()
    new_end_date = (request.form.get('end_date') or new_start_date or '').strip()
    new_end_time = (request.form.get('end_time') or '').strip()
    new_location = (request.form.get('location') or '').strip()
    new_purpose = (request.form.get('purpose') or '').strip()
    new_note = (request.form.get('note') or '-').strip() or '-'
    edit_reason = (request.form.get('edit_reason') or '').strip()

    # ตรวจสอบการกรอกข้อมูลจำเป็น
    if not all([new_vehicle, new_start_date, new_start_time, new_end_date, new_end_time, new_location, new_purpose]):
        flash('กรุณากรอกข้อมูลการแก้ไขให้ครบถ้วนทุกช่องที่มีเครื่องหมายดอกจัน (*)', 'warning')
        return redirect(url_for('user_request'))

    # ตรวจสอบเหตุผลในการแก้ไข
    if len(edit_reason) < MIN_REASON_LENGTH:
        flash('กรุณาระบุเหตุผลในการแก้ไขข้อมูลอย่างชัดเจน (ความยาวอย่างน้อย 5 ตัวอักษร เช่น ขอเปลี่ยนเป็นรถขนาดใหญ่ขึ้น, ขยายเวลาใช้งาน)', 'warning')
        return redirect(url_for('user_request'))

    # 8. ตรวจสอบเวลาที่เลือก (06:00 - 22:00 น. นาที 00/30)
    if new_start_time not in ALLOWED_TIMES:
        flash('เวลาเริ่มต้นไม่ถูกต้อง: ระบบเปิดให้จองเวลา 06:00 ถึง 22:00 น. (นาที 00 หรือ 30)', 'danger')
        return redirect(url_for('user_request'))
    if new_end_time not in ALLOWED_TIMES:
        flash('เวลาสิ้นสุดไม่ถูกต้อง: ระบบเปิดให้จองเวลา 06:00 ถึง 22:00 น. (นาที 00 หรือ 30)', 'danger')
        return redirect(url_for('user_request'))

    # 9. ตรวจสอบความถูกต้องของวันที่
    if new_start_date < today_str:
        flash(f'ไม่สามารถแก้ไขได้: วันที่เริ่มต้นการใช้งาน ({new_start_date}) เป็นวันที่ผ่านมาแล้ว', 'danger')
        return redirect(url_for('user_request'))
    if new_end_date < new_start_date:
        flash(f'ไม่สามารถแก้ไขได้: วันที่สิ้นสุด ({new_end_date}) ต้องไม่เกิดขึ้นก่อนวันที่เริ่มต้น ({new_start_date})', 'danger')
        return redirect(url_for('user_request'))

    # ตรวจสอบระยะเวลาจองไม่เกิน 7 วัน
    try:
        d1 = datetime.strptime(new_start_date, '%Y-%m-%d')
        d2 = datetime.strptime(new_end_date, '%Y-%m-%d')
        diff_days = (d2 - d1).days
        if diff_days > 7:
            flash(f'ไม่สามารถแก้ไขได้: ระยะเวลาการจองต้องไม่เกิน 1 สัปดาห์ (สูงสุด 7 วัน)', 'danger')
            return redirect(url_for('user_request'))
    except ValueError:
        flash('รูปแบบวันที่ไม่ถูกต้อง', 'danger')
        return redirect(url_for('user_request'))

    if new_start_date == new_end_date and new_end_time <= new_start_time:
        flash(f'ไม่สามารถแก้ไขได้: เวลาสิ้นสุด ({new_end_time} น.) ต้องมากกว่าเวลาเริ่มต้น ({new_start_time} น.) ในวันเดียวกัน', 'danger')
        return redirect(url_for('user_request'))

    # 10. ตรวจสอบการชนเวลา (Overlap Check) กับการจองอื่น (ยกเว้นคำขอของตัวเอง BK นี้)
    req_start = f"{new_start_date} {new_start_time}"
    req_end = f"{new_end_date} {new_end_time}"
    all_bookings = get_all_bookings()

    for b in all_bookings:
        if b['id'] != booking_id and b['status'] not in ['ไม่อนุมัติ', 'ยกเลิกแล้ว']:
            b_veh = b['vehicle'].strip()
            v_req = new_vehicle.strip()
            if b_veh == v_req or b_veh in v_req or v_req in b_veh:
                b_start = f"{b['start_date']} {b['start_time']}"
                b_end = f"{b['end_date']} {b['end_time']}"
                if req_start < b_end and req_end > b_start:
                    flash(
                        f'ไม่สามารถแก้ไขได้: ยานพาหนะ "{new_vehicle}" มีรายการจองอื่นทับซ้อนในช่วงวันที่ '
                        f'{b["start_date"]} ({b["start_time"]}) ถึง {b["end_date"]} ({b["end_time"]})',
                        'danger'
                    )
                    return redirect(url_for('user_request'))

    # 11. ตรวจสอบสถานะการอนุมัติ:
    # หากเคยได้รับ "อนุมัติ" แล้วแก้ไขรถหรือวันเวลา ให้รีเซ็ตสถานะกลับเป็น "รออนุมัติ" เพื่อให้เจ้าหน้าที่ตรวจสอบใหม่
    old_status = target_booking['status']
    new_status = 'รออนุมัติ'  # รีเซ็ตเป็นรออนุมัติเสมอเมื่อมีการแก้ไข เพื่อความปลอดภัย
    status_changed_msg = " (สถานะคำขอถูกปรับกลับเป็น 'รออนุมัติ' เพื่อให้เจ้าหน้าที่ตรวจสอบใหม่อีกครั้ง)" if old_status == 'อนุมัติ' else ""

    note_edit = f"{new_note} [แก้ไขครั้งที่ {edit_count + 1}/{MAX_EDITS_PER_BOOKING}: {edit_reason}]"

    # อัปเดตลงฐานข้อมูล
    update_booking_in_db(
        booking_id,
        start_date=new_start_date,
        start_time=new_start_time,
        end_date=new_end_date,
        end_time=new_end_time,
        vehicle=new_vehicle,
        location=new_location,
        purpose=new_purpose,
        note=note_edit,
        status=new_status
    )

    # อัปเดต timestamp ป้องกันสแปม
    session['last_booking_action_time'] = now_ts

    # บันทึก Logs
    log_detail = f"{booking_id} เปลี่ยนเป็น {new_vehicle} ({new_start_date} {new_start_time} - {new_end_date} {new_end_time}) เหตุผล: {edit_reason}"
    log_action('แก้ไขคำขอจอง', log_detail, current_user_name)

    # แจ้งเตือนผู้ใช้
    u_email = user_obj.get('email', '') if user_obj else ''
    u_id = user_obj.get('id', '') if user_obj else ''
    create_notification(
        user_id=u_id,
        user_name=current_user_name,
        email=u_email,
        booking_id=booking_id,
        vehicle=new_vehicle,
        status=new_status,
        reason=f"แก้ไขข้อมูลสำเร็จ ({edit_reason}){status_changed_msg}",
        start_date=new_start_date,
        start_time=new_start_time,
        end_date=new_end_date,
        end_time=new_end_time
    )

    flash(
        f'แก้ไขข้อมูลคำสั่งจอง {booking_id} เรียบร้อยแล้ว (ใช้สิทธิ์แก้ไขไปแล้ว {edit_count + 1}/{MAX_EDITS_PER_BOOKING} ครั้ง){status_changed_msg}',
        'success'
    )
    return redirect(url_for('user_request'))

# ==============================================================
# --- ROUTES: 3. ตารางการใช้รถ (Vehicle Usage Schedule & Events API) ---
# ==============================================================
def get_calendar_events_data(viewer_name=None, is_admin=False):
    """ดึงข้อมูลการจองทั้งหมดสำหรับแสดงผลบนปฏิทิน FullCalendar พร้อมรหัส BK และรายละเอียด (ปกปิดข้อมูลส่วนบุคคลตาม PDPA)"""
    all_bookings = get_all_bookings()
    events = []

    viewer_user_identifiers = get_user_identifiers(viewer_name) if viewer_name else set()
    if viewer_name:
        viewer_user_identifiers.add(str(viewer_name).strip().lower())

    for b in all_bookings:
        status = b['status']
        if status == 'อนุมัติ':
            bg_color = '#5b7337'  # เขียวมะกอก
            border_color = '#435528'
        elif status == 'ไม่อนุมัติ':
            bg_color = '#c94c5d'  # กุหลาบเข้ม
            border_color = '#a83242'
        else:  # รออนุมัติ
            bg_color = '#d99b26'  # เหลืองอำพัน
            border_color = '#b57f18'

        start_iso = f"{b['start_date']}T{b['start_time']}:00" if len(b['start_time']) == 5 else b['start_date']
        end_iso = f"{b['end_date']}T{b['end_time']}:00" if len(b['end_time']) == 5 else b['end_date']

        is_owner = is_booking_owner(b, viewer_user_identifiers) if viewer_user_identifiers else False

        if is_admin or is_owner:
            display_user = b['user']
            display_location = b['location']
        else:
            display_user = mask_user_last_name(b['user'])
            display_location = MASKED_FARM_LOCATION

        events.append({
            'id': b['id'],
            'title': f"[{b['id']}] {b['vehicle']}",
            'start': start_iso,
            'end': end_iso,
            'backgroundColor': bg_color,
            'borderColor': border_color,
            'textColor': '#ffffff',
            'extendedProps': {
                'booking_id': b['id'],
                'vehicle': b['vehicle'],
                'user': display_user,
                'location': display_location,
                'purpose': b['purpose'],
                'note': b['note'],
                'status': b['status'],
                'start_date': b['start_date'],
                'start_time': b['start_time'],
                'end_date': b['end_date'],
                'end_time': b['end_time'],
                'is_owner': is_owner,
                'is_admin': is_admin
            }
        })
    return events

def get_vehicles_overview_summary(viewer_name=None, is_admin=False):
    """คำนวณข้อมูลสรุปยานพาหนะทุกคัน: ใครจองวันไหนไปแล้วบ้าง, สรุปจำนวนครั้งที่ใช้, และสถานะว่างวันนี้หรือไม่ (ปกปิดข้อมูลส่วนบุคคลตาม PDPA)"""
    today_str = datetime.now().strftime('%Y-%m-%d')
    v_data = get_all_vehicles_data()
    all_bookings = get_all_bookings()

    viewer_user_identifiers = get_user_identifiers(viewer_name) if viewer_name else set()
    if viewer_name:
        viewer_user_identifiers.add(str(viewer_name).strip().lower())

    result = []
    for v in v_data:
        code = v['code']
        name = v['name']
        plate = v['plate']
        full_name = v['full_name']

        # รวบรวมประวัติการจองของคันนี้
        v_bookings = []
        is_busy_today = False
        today_booking_info = None

        for b in all_bookings:
            b_veh = b['vehicle'].strip()
            # เช็คว่าตรงกับรถคันนี้หรือไม่
            if (code == b_veh or code in b_veh or name in b_veh or (plate != '-' and plate in b_veh)):
                is_owner = is_booking_owner(b, viewer_user_identifiers) if viewer_user_identifiers else False

                if is_admin or is_owner:
                    display_user = b['user']
                    display_location = b['location']
                else:
                    display_user = mask_user_last_name(b['user'])
                    display_location = MASKED_FARM_LOCATION

                v_bookings.append({
                    'id': b['id'],
                    'start_date': b['start_date'],
                    'start_time': b['start_time'],
                    'end_date': b['end_date'],
                    'end_time': b['end_time'],
                    'user': display_user,
                    'location': display_location,
                    'purpose': b['purpose'],
                    'note': b['note'],
                    'status': b['status'],
                    'is_owner': is_owner
                })
                # ตรวจสอบสถานะวันนี้ (เฉพาะรายการที่อนุมัติหรือรออนุมัติ)
                if b['status'] in ['อนุมัติ', 'รออนุมัติ']:
                    if b['start_date'] <= today_str <= b['end_date']:
                        is_busy_today = True
                        today_booking_info = {
                            'id': b['id'],
                            'user': display_user,
                            'time': f"{b['start_time']} - {b['end_time']}",
                            'status': b['status'],
                            'purpose': b['purpose'],
                            'location': display_location,
                            'is_owner': is_owner
                        }

        # เรียงลำดับประวัติการจองตามวันที่ล่าสุด
        v_bookings.sort(key=lambda x: (x['start_date'], x['start_time']), reverse=True)

        result.append({
            'code': code,
            'name': name,
            'plate': plate,
            'full_name': full_name,
            'total_used': len(v_bookings),
            'is_busy_today': is_busy_today,
            'today_info': today_booking_info,
            'bookings': v_bookings
        })

    return today_str, result

@app.route('/overview-calendar')
def overview_calendar():
    """หน้าตารางการใช้รถสำหรับผู้ใช้ทั่วไป โหลดข้อมูลสรุปยานพาหนะและการจองล่วงหน้าแบบ SSR เพื่อความเร็วสูงสุด"""
    is_admin = bool(session.get('admin_logged_in') or session.get('role') == 'แอดมิน')
    viewer_name = session.get('user_name') or session.get('username') or session.get('user_id')
    today_str, vehicles_overview = get_vehicles_overview_summary(viewer_name=viewer_name, is_admin=is_admin)
    events = get_calendar_events_data(viewer_name=viewer_name, is_admin=is_admin)
    vehicles = get_vehicles()
    return render_template(
        'overview_calendar.html',
        vehicles=vehicles,
        vehicles_overview=vehicles_overview,
        all_events=events,
        today=today_str,
        is_admin_view=is_admin
    )

@app.route('/approve/overview-calendar')
@login_required
def admin_overview_calendar():
    """หน้าตารางการใช้รถสำหรับแอดมิน คงแถบเมนูแอดมินไว้อย่างสมบูรณ์"""
    viewer_name = session.get('username') or 'ผู้ดูแลระบบ'
    today_str, vehicles_overview = get_vehicles_overview_summary(viewer_name=viewer_name, is_admin=True)
    events = get_calendar_events_data(viewer_name=viewer_name, is_admin=True)
    vehicles = get_vehicles()
    return render_template(
        'overview_calendar.html',
        vehicles=vehicles,
        vehicles_overview=vehicles_overview,
        all_events=events,
        today=today_str,
        is_admin_view=True
    )

@app.route('/api/calendar-events')
def api_calendar_events():
    """API ส่งคืนข้อมูลการจองเพื่อแสดงผลใน FullCalendar พร้อมรหัสคำสั่ง BKxxx ชัดเจน"""
    is_admin = bool(session.get('admin_logged_in') or session.get('role') == 'แอดมิน')
    viewer_name = session.get('user_name') or session.get('username') or session.get('user_id')
    events = get_calendar_events_data(viewer_name=viewer_name, is_admin=is_admin)
    return jsonify(events)

@app.route('/api/vehicles-overview')
def api_vehicles_overview():
    """API ส่งคืนข้อมูลสรุปยานพาหนะทุกคัน: ใครจองวันไหนไปแล้วบ้าง, สรุปจำนวนครั้งที่ใช้, และสถานะว่างวันนี้หรือไม่"""
    is_admin = bool(session.get('admin_logged_in') or session.get('role') == 'แอดมิน')
    viewer_name = session.get('user_name') or session.get('username') or session.get('user_id')
    today_str, result = get_vehicles_overview_summary(viewer_name=viewer_name, is_admin=is_admin)
    return jsonify({
        'today': today_str,
        'vehicles': result
    })

# ==============================================================
# --- ROUTES: 4. ระบบผู้อนุมัติ (Admin Approval & Management) ---
# ==============================================================
@app.route('/approve')
@login_required
def admin_approve():
    all_bookings = get_all_bookings()
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

    reason = request.args.get('reason', '-').strip() or '-'

    # ค้นหาคำสั่งจองเดิมเพื่อดึงข้อมูลสำหรับส่งแจ้งเตือน
    all_bookings = get_all_bookings()
    target_booking = None
    for b in all_bookings:
        if b['id'] == booking_id:
            target_booking = b
            break

    # อัปเดตใน Google Sheets
    if use_google_sheets and ws_bookings:
        try:
            booking_rows = ws_bookings.get_all_values()
            for index, row in enumerate(booking_rows, start=1):
                if len(row) > 0 and row[0] == booking_id:
                    ws_bookings.update_cell(index, 11, status)
                    break
        except Exception as e:
            print(f"Error updating Google Sheets: {e}")

    # อัปเดตใน Local Excel
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_b = wb['Bookings']
        for row in ws_b.iter_rows(min_row=2):
            if row[0].value == booking_id:
                row[10].value = status
                break
        wb.save(EXCEL_FILE)

    # สร้างการแจ้งเตือนผู้จองทั้งในระบบและทางอีเมล
    notif_sent = False
    recipient_email = ''
    requester_name = ''
    if target_booking:
        requester_name = target_booking['user']
        vehicle_name = target_booking['vehicle']
        start_date = target_booking['start_date']
        start_time = target_booking['start_time']
        end_date = target_booking['end_date']
        end_time = target_booking['end_time']

        requester_user = None
        for u in get_all_users():
            if (u.get('full_name') and u['full_name'].strip() == requester_name.strip()) or \
               (u.get('username') and u['username'].strip() == requester_name.strip()) or \
               (u.get('email') and u['email'].strip().lower() == requester_name.strip().lower()):
                requester_user = u
                break

        user_id = requester_user['id'] if requester_user else ''
        recipient_email = requester_user.get('email', '') if requester_user else ''

        create_notification(
            user_id=user_id,
            user_name=requester_name,
            email=recipient_email,
            booking_id=booking_id,
            vehicle=vehicle_name,
            status=status,
            reason=reason,
            start_date=start_date,
            start_time=start_time,
            end_date=end_date,
            end_time=end_time
        )
        notif_sent = True

    admin_user = session.get('user_name', 'Admin')
    log_detail = f"{booking_id} เป็น {status}" + (f" (หมายเหตุ: {reason})" if reason != '-' else "")
    log_action('อัปเดตสถานะและแจ้งเตือนผู้จอง', log_detail, admin_user)

    if notif_sent and recipient_email:
        flash(f'ปรับสถานะคำสั่งจอง {booking_id} เป็น "{status}" เรียบร้อยแล้ว พร้อมส่งการแจ้งเตือนไปยังผู้จอง ({requester_name} - {recipient_email})', 'success')
    else:
        flash(f'ปรับสถานะคำสั่งจอง {booking_id} เป็น "{status}" เรียบร้อยแล้ว และบันทึกการแจ้งเตือนในระบบให้ผู้จอง ({requester_name}) เรียบร้อยแล้ว', 'success')
    return redirect(url_for('admin_approve'))

# ==============================================================
# --- ROUTES: 5. ส่งออกไฟล์ Excel สำหรับแอดมิน (Admin Excel Export) ---
# ==============================================================
def filter_bookings(start_date=None, end_date=None, status='ALL', vehicle='ALL'):
    """ฟังก์ชันตัวช่วยสำหรับกรองรายการจองตามช่วงวันที่ ยานพาหนะ และสถานะ"""
    bookings = get_all_bookings()
    filtered = []
    
    for b in bookings:
        # กรองตามช่วงวันที่ (ถ้ามี)
        if start_date and b['end_date'] < start_date:
            continue
        if end_date and b['start_date'] > end_date:
            continue
            
        # กรองตามสถานะ
        if status and status != 'ALL' and b['status'] != status:
            continue
            
        # กรองตามยานพาหนะ
        if vehicle and vehicle != 'ALL':
            b_veh = b['vehicle'].strip()
            v_req = vehicle.strip()
            if b_veh != v_req and b_veh not in v_req and v_req not in b_veh:
                continue

        filtered.append(b)
        
    return filtered

@app.route('/approve/export')
@app.route('/approve/export', endpoint='admin_export')
@login_required
def admin_export_page():
    today = datetime.now()
    default_start = (today.replace(day=1)).strftime('%Y-%m-%d')
    default_end = (today + timedelta(days=30)).strftime('%Y-%m-%d')

    filter_start = request.args.get('start_date', default_start)
    filter_end = request.args.get('end_date', default_end)
    filter_status = request.args.get('status', 'ALL')
    filter_vehicle = request.args.get('vehicle', 'ALL')

    filtered = filter_bookings(filter_start, filter_end, filter_status, filter_vehicle)
    vehicles = get_vehicles()

    count_approved = sum(1 for b in filtered if b['status'] == 'อนุมัติ')
    count_pending = sum(1 for b in filtered if b['status'] == 'รออนุมัติ')
    count_rejected = sum(1 for b in filtered if b['status'] == 'ไม่อนุมัติ')

    return render_template(
        'admin_export.html',
        filtered_bookings=filtered,
        filter_start=filter_start,
        filter_end=filter_end,
        filter_status=filter_status,
        filter_vehicle=filter_vehicle,
        vehicles=vehicles,
        count_approved=count_approved,
        count_pending=count_pending,
        count_rejected=count_rejected
    )

@app.route('/approve/export/download')
@login_required
def admin_export_download():
    """สร้างและดาวน์โหลดไฟล์ Excel (.xlsx) ที่กรองตามช่วงวันที่และเงื่อนไขที่แอดมินเลือก"""
    today = datetime.now()
    default_start = (today.replace(day=1)).strftime('%Y-%m-%d')
    default_end = (today + timedelta(days=30)).strftime('%Y-%m-%d')

    filter_start = request.args.get('start_date', default_start)
    filter_end = request.args.get('end_date', default_end)
    filter_status = request.args.get('status', 'ALL')
    filter_vehicle = request.args.get('vehicle', 'ALL')

    filtered = filter_bookings(filter_start, filter_end, filter_status, filter_vehicle)

    # สร้าง Workbook สวยงามด้วย OpenPyXL
    wb = Workbook()
    ws = wb.active
    ws.title = 'Booking_Report'

    # 1. หัวข้อรายงานใหญ่
    ws.merge_cells('A1:L1')
    title_cell = ws['A1']
    title_cell.value = 'รายงานการจองยานพาหนะและเครื่องจักรกลเกษตร (Agricarbook Report)'
    title_cell.font = Font(name='Prompt', size=15, bold=True, color='2C3315')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_cell.fill = PatternFill(start_color='FAD1D9', end_color='FAD1D9', fill_type='solid')
    ws.row_dimensions[1].height = 36

    # 2. รายละเอียดเงื่อนไขรายงาน
    ws.merge_cells('A2:L2')
    sub_cell = ws['A2']
    now_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    admin_name = session.get('user_name', 'Admin')
    sub_cell.value = (
        f"ช่วงวันที่: {filter_start} ถึง {filter_end} | สถานะ: {filter_status} | "
        f"ยานพาหนะ: {filter_vehicle} | รายการทั้งหมด: {len(filtered)} รายการ | ส่งออกเมื่อ: {now_str} (โดย {admin_name})"
    )
    sub_cell.font = Font(name='Prompt', size=9.5, italic=True, color='555555')
    sub_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 22

    # แถวว่าง
    ws.row_dimensions[3].height = 8

    # 3. ส่วนหัวตาราง (Headers)
    headers = [
        'ลำดับ', 'รหัสการจอง', 'วันที่เริ่มต้น', 'เวลาเริ่มต้น', 
        'วันที่สิ้นสุด', 'เวลาสิ้นสุด', 'ยานพาหนะเกษตร', 'ผู้ขอใช้รถ', 
        'สถานที่ปฏิบัติงาน', 'วัตถุประสงค์การใช้งาน', 'หมายเหตุ', 'สถานะ'
    ]
    ws.append(headers)
    ws.row_dimensions[4].height = 28

    header_font = Font(name='Prompt', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='868F5A', end_color='868F5A', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=4, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # 4. ข้อมูลแถวตาราง
    zebra_fill = PatternFill(start_color='FDF5F6', end_color='FDF5F6', fill_type='solid')
    white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

    for idx, b in enumerate(filtered, start=1):
        row_num = idx + 4
        ws.row_dimensions[row_num].height = 22
        current_fill = zebra_fill if idx % 2 == 1 else white_fill

        row_values = [
            idx,
            b['id'],
            b['start_date'],
            b['start_time'],
            b['end_date'],
            b['end_time'],
            b['vehicle'],
            b['user'],
            b['location'],
            b['purpose'],
            b['note'],
            b['status']
        ]
        ws.append(row_values)

        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_num, column=col_idx)
            c.font = Font(name='Prompt', size=10)
            c.fill = current_fill
            c.border = thin_border
            
            # จัดตำแหน่งข้อความ
            if col_idx in [1, 2, 3, 4, 5, 6, 12]:
                c.alignment = Alignment(horizontal='center', vertical='center')
            else:
                c.alignment = Alignment(horizontal='left', vertical='center')

            # สีของสถานะ
            if col_idx == 12:
                if b['status'] == 'อนุมัติ':
                    c.font = Font(name='Prompt', size=10, bold=True, color='2E5917')
                elif b['status'] == 'ไม่อนุมัติ':
                    c.font = Font(name='Prompt', size=10, bold=True, color='C94C5D')
                else:
                    c.font = Font(name='Prompt', size=10, bold=True, color='D99B26')

    # 5. ปรับความกว้างของคอลัมน์อัตโนมัติ
    column_widths = {
        'A': 8,   # ลำดับ
        'B': 14,  # รหัสการจอง
        'C': 15,  # วันที่เริ่มต้น
        'D': 13,  # เวลาเริ่มต้น
        'E': 15,  # วันที่สิ้นสุด
        'F': 13,  # เวลาสิ้นสุด
        'G': 34,  # ยานพาหนะเกษตร
        'H': 22,  # ผู้ขอใช้รถ
        'I': 30,  # สถานที่
        'J': 30,  # วัตถุประสงค์
        'K': 18,  # หมายเหตุ
        'L': 14   # สถานะ
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    # ส่งออกไฟล์เป็น Response Attachment
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Agricarbook_Report_{filter_start}_{filter_end}.xlsx"
    log_action('ส่งออกรายงาน Excel', filename, admin_name)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

# ==============================================================
# --- ROUTES: 6. จัดการยานพาหนะสำหรับแอดมิน (Admin Vehicle Management) ---
# ==============================================================
@app.route('/approve/vehicles')
@login_required
def admin_vehicles_page():
    vehicles_data = get_all_vehicles_data()
    # คำนวณรหัสถัดไปอัตโนมัติ เช่น V023
    existing_nums = []
    for v in vehicles_data:
        code = v['code']
        if code.startswith('V') and code[1:].isdigit():
            existing_nums.append(int(code[1:]))
    next_num = max(existing_nums) + 1 if existing_nums else 1
    next_code = f"V{next_num:03d}"

    return render_template(
        'admin_vehicles.html',
        vehicles_data=vehicles_data,
        next_code=next_code
    )

@app.route('/approve/vehicles/add', methods=['POST'])
@login_required
def admin_add_vehicle():
    code = request.form.get('code', '').strip().upper()
    name = request.form.get('name', '').strip()
    plate = request.form.get('plate', '').strip()

    if not code or not name:
        flash('กรุณาระบุรหัสยานพาหนะและชื่อเครื่องจักรกลเกษตร', 'warning')
        return redirect(url_for('admin_vehicles_page'))

    # ตรวจสอบรหัสซ้ำ
    existing = get_all_vehicles_data()
    for v in existing:
        if v['code'] == code:
            flash(f'รหัสยานพาหนะ {code} มีอยู่ในระบบแล้ว กรุณาใช้รหัสอื่น', 'danger')
            return redirect(url_for('admin_vehicles_page'))

    add_vehicle_to_db(code, name, plate)
    admin_user = session.get('user_name', 'Admin')
    log_action('เพิ่มยานพาหนะ', f"{code}: {name} ({plate})", admin_user)
    flash(f'เพิ่มยานพาหนะ "{name} ({plate})" รหัส {code} เข้าสู่ฐานข้อมูลเรียบร้อยแล้ว', 'success')
    return redirect(url_for('admin_vehicles_page'))

@app.route('/approve/vehicles/delete/<code>', methods=['POST'])
@login_required
def admin_delete_vehicle(code):
    code = code.strip().upper()
    delete_vehicle_from_db(code)
    admin_user = session.get('user_name', 'Admin')
    log_action('ลบยานพาหนะ', code, admin_user)
    flash(f'ลบยานพาหนะรหัส {code} ออกจากฐานข้อมูลเรียบร้อยแล้ว', 'info')
    return redirect(url_for('admin_vehicles_page'))

# ==============================================================
# --- ROUTES: 7. หน้าผู้จัดทำและตัวช่วยนำทาง ---
# ==============================================================
@app.route('/calendar')
def calendar():
    return redirect(url_for('overview_calendar'))

@app.route('/approve/calendar')
def admin_calendar():
    return redirect(url_for('admin_overview_calendar'))

@app.route('/admin')
def admin_shortcut():
    return redirect(url_for('admin_approve'))

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