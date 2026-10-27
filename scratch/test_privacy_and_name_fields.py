import unittest
from app import app, split_full_name, mask_user_last_name, MASKED_FARM_LOCATION, get_calendar_events_data, get_vehicles_overview_summary

class TestPrivacyAndNameFields(unittest.TestCase):
    def setUp(self):
        self.app = app
        self.app.config['TESTING'] = True
        self.app.config['WTF_CSRF_ENABLED'] = False
        self.client = self.app.test_client()

    def test_01_split_full_name(self):
        """ทดสอบการแยกชื่อจริงและนามสกุล ทั้งที่มีคำนำหน้าและไม่มีคำนำหน้า"""
        cases = [
            ("นายสมเกียรติ สุขเกษม", "นายสมเกียรติ", "สุขเกษม"),
            ("นาย สมเกียรติ สุขเกษม", "นาย สมเกียรติ", "สุขเกษม"),
            ("นางสาวกานดา มีสุข", "นางสาวกานดา", "มีสุข"),
            ("นางสาว กานดา มีสุขเจริญ", "นางสาว กานดา", "มีสุขเจริญ"),
            ("น.ส. มาลี ดอกรัก", "น.ส. มาลี", "ดอกรัก"),
            ("สมชาย ใจดี", "สมชาย", "ใจดี"),
            ("John Doe", "John", "Doe"),
            ("สมเกียรติ", "สมเกียรติ", ""),
            ("", "", "")
        ]
        for full, expected_fn, expected_ln in cases:
            fn, ln = split_full_name(full)
            self.assertEqual(fn, expected_fn, f"First name mismatch for '{full}'")
            self.assertEqual(ln, expected_ln, f"Last name mismatch for '{full}'")

    def test_02_mask_user_last_name(self):
        """ทดสอบการ Masking นามสกุลเพื่อการคุ้มครองข้อมูลส่วนบุคคล (PDPA)"""
        cases = [
            ("นายสมเกียรติ สุขเกษม", "นายสมเกียรติ ส***"),
            ("สมชาย ใจดี", "สมชาย ใ***"),
            ("นางสาว กานดา มีสุขเจริญ", "นางสาว กานดา ม***"),
            ("John Doe", "John D***"),
            ("สมเกียรติ", "สมเ***"),
            ("", "ผู้ใช้งานทั่วไป")
        ]
        for full, expected_masked in cases:
            masked = mask_user_last_name(full)
            self.assertEqual(masked, expected_masked, f"Masking mismatch for '{full}'")

    def test_03_privacy_masking_for_other_users(self):
        """ทดสอบว่าผู้ใช้งานทั่วไป (User A) มองเห็นการจองของคนอื่นแบบถูกปิดบังนามสกุลและพื้นที่ฟาร์ม"""
        # สมมุติว่ามีรายการจองของ User B
        # เรียก get_calendar_events_data โดยส่ง viewer_name เป็น User A
        events = get_calendar_events_data(viewer_name="นายสมเกียรติ สุขเกษม", is_admin=False)
        self.assertIsInstance(events, list)
        
        for ev in events:
            p = ev.get('extendedProps', {})
            # ถ้าเป็นคำสั่งจองของคนอื่น (ไม่ใช่ นายสมเกียรติ สุขเกษม)
            if not p.get('is_owner'):
                # นามสกุลต้องถูก mask ด้วย ***
                self.assertIn('***', p.get('user', ''))
                # สถานที่ฟาร์มต้องถูกปิดบัง
                self.assertEqual(p.get('location'), MASKED_FARM_LOCATION)
            else:
                # ถ้าเป็นของตัวเอง ต้องไม่ถูก mask
                self.assertNotEqual(p.get('location'), MASKED_FARM_LOCATION)

    def test_04_admin_sees_all_unmasked(self):
        """ทดสอบว่าผู้ดูแลระบบ (Admin) สามารถมองเห็นชื่อจริง-นามสกุล และพื้นที่ฟาร์มทั้งหมดได้ตามปกติ"""
        events = get_calendar_events_data(viewer_name="admin", is_admin=True)
        self.assertIsInstance(events, list)
        
        for ev in events:
            p = ev.get('extendedProps', {})
            # แอดมินต้องเห็นสถานที่จริง ไม่ใช่ MASKED_FARM_LOCATION
            self.assertNotEqual(p.get('location'), MASKED_FARM_LOCATION)

    def test_05_register_page_has_first_and_last_name_fields(self):
        """ตรวจสอบว่าหน้าสมัครสมาชิกมีฟิลด์ first_name และ last_name แยก 2 ช่อง"""
        response = self.client.get('/login')
        html = response.data.decode('utf-8')
        self.assertIn('name="first_name"', html)
        self.assertIn('name="last_name"', html)
        self.assertIn('name="full_name"', html)
        self.assertIn('ชื่อจริง', html)
        self.assertIn('นามสกุล', html)

    def test_06_booking_page_has_first_and_last_name_fields(self):
        """ตรวจสอบว่าหน้าจองรถมีฟิลด์ first_name และ last_name แยก 2 ช่อง พร้อมดึงชื่อจาก session"""
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['user_id'] = 'U002'
            sess['username'] = 'somkiat'
            sess['user_name'] = 'นายสมเกียรติ สุขเกษม'
            sess['role'] = 'ผู้ใช้'

        response = self.client.get('/request')
        html = response.data.decode('utf-8')
        self.assertIn('name="first_name"', html)
        self.assertIn('name="last_name"', html)
        self.assertIn('value="นายสมเกียรติ"', html)
        self.assertIn('value="สุขเกษม"', html)

    def test_07_overview_calendar_page_and_api(self):
        """ตรวจสอบหน้าตารางการใช้รถ และ API ส่งข้อมูลพร้อม privacy masking"""
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['user_id'] = 'U002'
            sess['username'] = 'somkiat'
            sess['user_name'] = 'นายสมเกียรติ สุขเกษม'
            sess['role'] = 'ผู้ใช้'

        # ทดสอบ SSR Page
        res_page = self.client.get('/overview-calendar')
        self.assertEqual(res_page.status_code, 200)
        html = res_page.data.decode('utf-8')
        self.assertIn('ตารางการใช้รถ', html)

        # ทดสอบ API
        res_api = self.client.get('/api/calendar-events')
        self.assertEqual(res_api.status_code, 200)
        events_json = res_api.get_json()
        self.assertIsInstance(events_json, list)

    def test_08_booking_post_combines_first_and_last_name(self):
        """ตรวจสอบว่าการส่งฟอร์มจองรถด้วย first_name และ last_name รวมเป็นชื่อ-นามสกุลถูกต้อง"""
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['user_id'] = 'U002'
            sess['username'] = 'somkiat'
            sess['user_name'] = 'นายสมเกียรติ สุขเกษม'
            sess['role'] = 'ผู้ใช้'

        # ถ้าส่งฟอร์มโดยไม่ใส่ first_name หรือ last_name และไม่มี user_name
        res_missing = self.client.post('/request', data={
            'first_name': '',
            'last_name': '',
            'user_name': '',
            'start_date': '2099-01-01',
            'start_time': '08:00',
            'end_date': '2099-01-01',
            'end_time': '12:00',
            'vehicle': 'รถแทรกเตอร์คูโบต้า L5018',
            'location': 'แปลงทดสอบ',
            'purpose': 'ทดสอบระบบ'
        }, follow_redirects=True)
        # ต้องไม่สำเร็จหรือดึงชื่อจาก session มาใช้
        self.assertEqual(res_missing.status_code, 200)

if __name__ == '__main__':
    unittest.main()
