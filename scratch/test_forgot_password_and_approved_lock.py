import unittest
from datetime import datetime
from app import (
    app, find_user_by_identifier, update_user_password, get_user_bookings,
    get_booking_by_id, update_booking_in_db, mask_email
)

class TestForgotPasswordAndApprovedLock(unittest.TestCase):
    def setUp(self):
        self.app = app
        self.app.config['TESTING'] = True
        self.app.config['WTF_CSRF_ENABLED'] = False
        self.client = self.app.test_client()

    def test_01_mask_email(self):
        """ทดสอบการปิดบังข้อมูลอีเมลเพื่อความปลอดภัย"""
        self.assertEqual(mask_email('somchai@farm.go.th'), 's***i@farm.go.th')
        self.assertEqual(mask_email('ab@test.com'), 'a***@test.com')
        self.assertEqual(mask_email(''), '')

    def test_02_forgot_password_send_otp_success(self):
        """ทดสอบการขอรับ OTP ลืมรหัสผ่านด้วยอีเมล และเลขบัตรประชาชน 13 หลัก"""
        # ด้วยอีเมลของนายสมชาย
        res_email = self.client.post('/api/forgot-password/send-otp', json={'identifier': 'somchai@farm.go.th'})
        self.assertEqual(res_email.status_code, 200)
        data_email = res_email.get_json()
        self.assertTrue(data_email['success'])
        self.assertIn('s***i@farm.go.th', data_email['masked_email'])
        self.assertIn('demo_otp', data_email)

        # ด้วยเลขประจำตัวประชาชนของนายสมชาย (1-1001-00111-11-1)
        res_cid = self.client.post('/api/forgot-password/send-otp', json={'identifier': '1-1001-00111-11-1'})
        self.assertEqual(res_cid.status_code, 200)
        data_cid = res_cid.get_json()
        self.assertTrue(data_cid['success'])

    def test_03_forgot_password_send_otp_not_found(self):
        """ทดสอบการขอ OTP กับบัญชีที่ไม่มีในระบบ"""
        res = self.client.post('/api/forgot-password/send-otp', json={'identifier': 'nonexistent@farm.com'})
        self.assertEqual(res.status_code, 404)
        data = res.get_json()
        self.assertFalse(data['success'])

    def test_04_forgot_password_verify_and_reset_flow(self):
        """ทดสอบขั้นตอนการยืนยัน OTP และตั้งรหัสผ่านใหม่ครบทั้ง Flow พร้อมล็อกอินด้วยรหัสใหม่"""
        with self.client:
            # 1. ขอ OTP
            res_otp = self.client.post('/api/forgot-password/send-otp', json={'identifier': 'somchai@farm.go.th'})
            self.assertEqual(res_otp.status_code, 200)
            data_otp = res_otp.get_json()
            demo_code = data_otp['demo_otp']

            # 2. ยืนยัน OTP ผิดพลาด
            res_wrong = self.client.post('/api/forgot-password/verify-otp', json={'otp': '000000'})
            self.assertEqual(res_wrong.status_code, 400)
            self.assertFalse(res_wrong.get_json()['success'])

            # 3. ยืนยัน OTP ถูกต้อง
            res_correct = self.client.post('/api/forgot-password/verify-otp', json={'otp': demo_code})
            self.assertEqual(res_correct.status_code, 200)
            self.assertTrue(res_correct.get_json()['success'])

            # 4. ตั้งรหัสผ่านใหม่ (เช่น NewPass#999)
            res_reset = self.client.post('/forgot-password/reset', data={
                'new_password': 'NewPass#999',
                'confirm_password': 'NewPass#999'
            }, follow_redirects=True)
            self.assertEqual(res_reset.status_code, 200)
            html = res_reset.data.decode('utf-8')
            self.assertIn('ตั้งรหัสผ่านใหม่สำเร็จเรียบร้อยแล้ว', html)

            # 5. ทดสอบเข้าสู่ระบบด้วยรหัสผ่านใหม่
            res_login_new = self.client.post('/login', data={
                'identifier': 'somchai@farm.go.th',
                'password': 'NewPass#999'
            }, follow_redirects=False)
            # ต้องเข้าสู่ระบบสำเร็จ (Redirect ไป /request)
            self.assertEqual(res_login_new.status_code, 302)
            self.assertIn('/request', res_login_new.location)

    def test_05_approved_booking_cannot_be_cancelled(self):
        """ทดสอบว่ารถที่ได้รับการอนุมัติแล้ว ไม่สามารถยกเลิกได้ทั้งบน UI และ Backend Route"""
        # สร้างรายการจองจำลองและกำหนดสถานะเป็น 'อนุมัติ'
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['user_id'] = 'U001'
            sess['username'] = '100123456789'
            sess['user_name'] = 'นายสมชาย ใจดี'
            sess['role'] = 'ผู้ใช้'

        # ดึงรายการจองของนายสมชาย
        my_b, _ = get_user_bookings('นายสมชาย ใจดี')
        self.assertTrue(len(my_b) > 0)
        target_bid = my_b[0]['id']

        # ตั้งสถานะเป็น อนุมัติ
        update_booking_in_db(target_bid, status='อนุมัติ', end_date='2099-12-31')

        # ตรวจสอบ get_user_bookings อีกครั้ง
        my_b_after, _ = get_user_bookings('นายสมชาย ใจดี')
        approved_b = next((b for b in my_b_after if b['id'] == target_bid), None)
        self.assertIsNotNone(approved_b)
        self.assertEqual(approved_b['status'], 'อนุมัติ')
        self.assertFalse(approved_b['can_cancel'], "Approved booking must not be cancellable")
        self.assertFalse(approved_b['can_edit'], "Approved booking must not be editable")

        # ส่งคำขอยกเลิกไปยัง Backend POST /booking/cancel/<id>
        res_cancel = self.client.post(f'/booking/cancel/{target_bid}', data={
            'cancel_reason': 'ขอเปลี่ยนใจยกเลิก'
        }, follow_redirects=True)
        html = res_cancel.data.decode('utf-8')
        self.assertIn('ได้รับการอนุมัติการจองแล้ว จึงไม่สามารถยกเลิกได้', html)

        # ตรวจสอบว่าในฐานข้อมูลสถานะยังคงเป็น อนุมัติ ไม่ถูกเปลี่ยนเป็น ยกเลิกแล้ว
        b_check = get_booking_by_id(target_bid)
        self.assertEqual(b_check['status'], 'อนุมัติ')

    def test_06_approved_booking_cannot_be_edited(self):
        """ทดสอบว่ารถที่ได้รับการอนุมัติแล้ว ไม่สามารถแก้ไขได้ทั้งบน UI และ Backend Route"""
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['user_id'] = 'U001'
            sess['username'] = '100123456789'
            sess['user_name'] = 'นายสมชาย ใจดี'
            sess['role'] = 'ผู้ใช้'

        my_b, _ = get_user_bookings('นายสมชาย ใจดี')
        target_bid = my_b[0]['id']
        update_booking_in_db(target_bid, status='อนุมัติ', end_date='2099-12-31')

        # ส่งคำขอแก้ไขไปยัง Backend POST /booking/edit/<id>
        res_edit = self.client.post(f'/booking/edit/{target_bid}', data={
            'start_date': '2099-11-01',
            'start_time': '08:00',
            'end_date': '2099-11-02',
            'end_time': '17:00',
            'location': 'แปลงใหม่',
            'purpose': 'เกี่ยวข้าวใหม่',
            'edit_reason': 'ขอเลื่อนวัน'
        }, follow_redirects=True)
        html = res_edit.data.decode('utf-8')
        self.assertIn('ได้รับการอนุมัติการจองแล้ว จึงไม่สามารถแก้ไขได้', html)

    def test_07_pending_booking_can_still_be_cancelled_and_edited(self):
        """ทดสอบว่ารถที่อยู่ในสถานะ 'รออนุมัติ' ยังคงสามารถแก้ไขและยกเลิกได้ตามปกติ"""
        my_b, _ = get_user_bookings('นายสมชาย ใจดี')
        target_bid = my_b[0]['id']
        update_booking_in_db(target_bid, status='รออนุมัติ', end_date='2099-12-31')

        my_b_after, _ = get_user_bookings('นายสมชาย ใจดี')
        pending_b = next((b for b in my_b_after if b['id'] == target_bid), None)
        self.assertIsNotNone(pending_b)
        self.assertEqual(pending_b['status'], 'รออนุมัติ')
        self.assertTrue(pending_b['can_cancel'], "Pending booking should be cancellable")
        self.assertTrue(pending_b['can_edit'], "Pending booking should be editable")

if __name__ == '__main__':
    unittest.main()
