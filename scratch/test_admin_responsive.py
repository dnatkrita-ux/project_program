import unittest
from app import app

class TestAdminResponsive(unittest.TestCase):
    def setUp(self):
        self.app = app
        self.app.config['TESTING'] = True
        self.client = self.app.test_client()

    def test_admin_page_responsive_elements(self):
        """ตรวจสอบว่าหน้า Admin มีโครงสร้างที่รองรับทุกขนาดหน้าจอ (Table View + Mobile Card View + Search + Filter Chips)"""
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['admin_logged_in'] = True
            sess['user_id'] = 'ADMIN01'
            sess['username'] = 'admin'
            sess['user_name'] = 'ผู้ดูแลระบบ (Admin)'
            sess['role'] = 'admin'

        res = self.client.get('/approve')
        self.assertEqual(res.status_code, 200)
        html = res.data.decode('utf-8')

        # ตรวจสอบหัวข้อประวัติการจอง
        self.assertIn('ประวัติรายการจองทั้งหมด', html)

        # ตรวจสอบช่องค้นหาแบบเรียลไทม์
        self.assertIn('id="adminHistorySearch"', html)

        # ตรวจสอบปุ่มชิปกรองสถานะ
        self.assertIn('data-status="ALL"', html)
        self.assertIn('data-status="รออนุมัติ"', html)
        self.assertIn('data-status="อนุมัติ"', html)
        self.assertIn('data-status="ไม่อนุมัติ"', html)

        # ตรวจสอบว่าใช้ตารางรูปแบบเดียวกันทุกอุปกรณ์ (ไม่มีการแยกเป็นโหมดการ์ดมือถือที่ทำให้ระบบต่างกัน)
        self.assertIn('id="adminTableView"', html)
        self.assertIn('id="adminBookingsTable"', html)
        self.assertNotIn('id="btnViewCards"', html)
        self.assertNotIn('id="adminCardView"', html)

        # ตรวจสอบแท็กและสไตล์ที่ป้องกันข้อความแตก/ตกบรรทัด
        self.assertIn('class="farm-zebra-table', html)
        self.assertIn('status-badge', html)

    def test_css_rules_prevent_vertical_status_badge(self):
        """ตรวจสอบว่าไฟล์ CSS มีกฎป้องกัน status-badge และ table column ไม่ให้บีบจนตัวอักษรแตกแถว"""
        with open('static/css/farm-theme.css', 'r', encoding='utf-8') as f:
            css = f.read()

        # status-badge ต้องมี white-space: nowrap !important
        self.assertIn('white-space: nowrap !important;', css)
        self.assertIn('word-break: keep-all !important;', css)

        # farm-zebra-table ต้องมี min-width
        self.assertIn('min-width: 880px;', css)

        # farm-table-card-yellow ต้องไม่มี max-width: 760px
        self.assertNotIn('.farm-table-card-yellow {\n    background: #ffffff;\n    border: 2px solid var(--farm-rose-border);\n    border-radius: var(--farm-card-radius);\n    box-shadow: 0 8px 24px rgba(134, 143, 90, 0.1);\n    overflow: hidden;\n    width: 100%;\n    max-width: 760px;', css)

    def test_booking_form_mobile_date_alignment_and_no_past_dates(self):
        """ตรวจสอบว่าแบบฟอร์มขอจองใช้โครงสร้าง row g-2 ทุกช่องทำให้ขนาดเท่ากันพอดี ไม่ล้นจอ และมี Flatpickr / การล็อควันย้อนหลัง"""
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['user_id'] = 'U001'
            sess['username'] = '100123456789'
            sess['user_name'] = 'นายสมชาย ใจดี'
            sess['role'] = 'ผู้ใช้'

        res = self.client.get('/request')
        self.assertEqual(res.status_code, 200)
        html = res.data.decode('utf-8')

        # 1. ตรวจสอบว่ามีลิงก์และสคริปต์ Flatpickr
        self.assertIn('flatpickr.min.css', html)
        self.assertIn('flatpickr', html)
        self.assertIn('th.js', html)

        # 2. ตรวจสอบการตั้งค่า disableMobile เพื่อบล็อกไม่ให้มือถือใช้ Native Picker ที่กดวันอดีตได้
        self.assertIn('disableMobile: true', html)
        self.assertIn('minDate: "today"', html)

        # 3. ตรวจสอบว่าช่องวันที่และเวลาถูกเปลี่ยนเป็น row g-2 (col-12 col-md-6) เหมือนกับช่องชื่อและนามสกุล
        self.assertIn('id="startDateInput"', html)
        self.assertIn('id="endDateInput"', html)
        self.assertIn('id="startTimeInput"', html)
        self.assertIn('id="endTimeInput"', html)
        self.assertIn('-- เลือกเวลาเริ่มต้น --', html)
        self.assertIn('-- เลือกเวลาสิ้นสุด --', html)

        # 4. ตรวจสอบว่าใน farm-theme.css มีการกำหนด CSS สไตล์ให้ขนาดเท่ากันพอดี และไม่ล้นจอ
        with open('static/css/farm-theme.css', 'r', encoding='utf-8') as f:
            css = f.read()

        self.assertIn('.flatpickr-calendar', css)
        self.assertIn('-webkit-appearance: none;', css)
        self.assertIn('width: 100% !important;', css)
        self.assertIn('box-sizing: border-box !important;', css)

    def test_week_view_readability_and_responsive_styling(self):
        """ตรวจสอบว่าหน้าตารางการใช้รถมีสไตล์และการตั้งค่ามุมมองสัปดาห์ที่อ่านง่ายในทุกอุปกรณ์ ไม่ตัดคำแนวตั้ง และตัดเวลา 00:00-05:00 ทิ้ง"""
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['admin_logged_in'] = True
            sess['username'] = 'admin'
            sess['role'] = 'admin'

        res = self.client.get('/approve/overview-calendar')
        self.assertEqual(res.status_code, 200)
        html = res.data.decode('utf-8')

        # 1. ตรวจสอบการตั้งค่าช่วงเวลาทำงาน (06:00 - 22:30) และตัด allDaySlot
        self.assertIn("slotMinTime: '06:00:00'", html)
        self.assertIn("slotMaxTime: '22:30:00'", html)
        self.assertIn("allDaySlot: false", html)

        # 2. ตรวจสอบ Header วันแบบชัดเจน (fc-week-day-header) ไม่แตกบรรทัด
        self.assertIn('fc-week-day-header', html)
        self.assertIn('fc-timegrid-card', html)

        # 3. ตรวจสอบการตั้งค่าความกว้างขั้นต่ำและการเลื่อนแนวนอนบนหน้าจอมือถือ
        self.assertIn('.fc-timeGridWeek-view .fc-scrollgrid', html)
        self.assertIn('min-width: 700px !important;', html)

if __name__ == '__main__':
    unittest.main()
