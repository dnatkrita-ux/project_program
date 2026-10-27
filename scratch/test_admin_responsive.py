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

        # ตรวจสอบปุ่มสลับมุมมอง (ตาราง vs การ์ดมือถือ)
        self.assertIn('id="btnViewTable"', html)
        self.assertIn('id="btnViewCards"', html)

        # ตรวจสอบคอนเทนเนอร์ทั้ง 2 มุมมอง
        self.assertIn('id="adminTableView"', html)
        self.assertIn('id="adminCardView"', html)

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

if __name__ == '__main__':
    unittest.main()
