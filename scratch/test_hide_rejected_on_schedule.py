import unittest
from app import (
    app, get_calendar_events_data, get_vehicles_overview_summary,
    get_user_rejected_notifications, create_notification
)

class TestHideRejectedOnSchedule(unittest.TestCase):
    def setUp(self):
        self.app = app
        self.app.config['TESTING'] = True
        self.client = self.app.test_client()

    def test_01_public_calendar_events_exclude_rejected(self):
        """ตรวจสอบว่า get_calendar_events_data สำหรับผู้ใช้ทั่วไป (is_admin=False) จะต้องไม่มีรายการ 'ไม่อนุมัติ'"""
        events = get_calendar_events_data(viewer_name=None, is_admin=False)
        self.assertTrue(len(events) > 0, "ควรมีรายการจองบนปฏิทิน")
        
        for ev in events:
            p = ev.get('extendedProps', {})
            status = p.get('status')
            self.assertNotEqual(
                status, 'ไม่อนุมัติ', 
                f"พบคำขอ {p.get('booking_id')} ที่มีสถานะ 'ไม่อนุมัติ' ปรากฏบนปฏิทินสาธารณะ!"
            )
            self.assertNotEqual(
                status, 'ยกเลิกแล้ว',
                f"พบคำขอ {p.get('booking_id')} ที่ถูกยกเลิกแล้วปรากฏบนปฏิทิน!"
            )

    def test_02_vehicles_overview_bookings_exclude_rejected_for_public(self):
        """ตรวจสอบว่า get_vehicles_overview_summary สำหรับผู้ใช้ทั่วไป (is_admin=False) ต้องไม่แสดงรายการ 'ไม่อนุมัติ' ในประวัติการใช้งานรถ"""
        today_str, v_summary = get_vehicles_overview_summary(viewer_name=None, is_admin=False)
        self.assertTrue(len(v_summary) > 0)

        for v in v_summary:
            for b in v['bookings']:
                self.assertNotEqual(
                    b['status'], 'ไม่อนุมัติ',
                    f"พบคำขอ {b['id']} ที่มีสถานะ 'ไม่อนุมัติ' ในตารางการจองของรถ {v['name']}!"
                )
                self.assertNotEqual(
                    b['status'], 'ยกเลิกแล้ว',
                    f"พบคำขอ {b['id']} ที่ถูกยกเลิกแล้วในตารางการจองของรถ {v['name']}!"
                )

    def test_03_overview_calendar_ui_for_general_user(self):
        """ตรวจสอบหน้า /overview-calendar สำหรับผู้ใช้ทั่วไปว่าไม่มีตัวเลือก 'ไม่อนุมัติ' ในตัวกรองและสัญลักษณ์สถานะ"""
        res = self.client.get('/overview-calendar')
        self.assertEqual(res.status_code, 200)
        html = res.get_data(as_text=True)

        # ตัวกรองสถานะต้องไม่มี <option value="ไม่อนุมัติ">
        self.assertNotIn('<option value="ไม่อนุมัติ">✖ ไม่อนุมัติ</option>', html)
        # สัญลักษณ์สถานะต้องไม่มี badge ไม่อนุมัติ
        self.assertNotIn('✖ ไม่อนุมัติ', html)

    def test_04_user_with_rejection_sees_private_notification(self):
        """ตรวจสอบว่าผู้ใช้ที่มีคำขอไม่ได้รับการอนุมัติ จะเห็นกล่องแจ้งเตือนเฉพาะตัวบนหน้าตารางการใช้รถ"""
        # สร้าง notification ทดสอบสำหรับนายสมชาย
        create_notification(
            user_id="U001",
            user_name="นายสมชาย ใจดี",
            email="somchai@farm.go.th",
            booking_id="BK999_TEST",
            vehicle="รถเกี่ยวนวดข้าว (นข 1234)",
            status="ไม่อนุมัติ",
            reason="ยานพาหนะอยู่ระหว่างการซ่อมบำรุงประจำปี"
        )

        # จำลองล็อกอินเป็นนายสมชาย
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['user_name'] = 'นายสมชาย ใจดี'
            sess['email'] = 'somchai@farm.go.th'
            sess['role'] = 'user'

        res = self.client.get('/overview-calendar')
        self.assertEqual(res.status_code, 200)
        html = res.get_data(as_text=True)

        # นายสมชายต้องเห็นการแจ้งเตือนเฉพาะตัว
        self.assertIn('แจ้งเตือนเฉพาะคุณ: รายการคำขอจองที่ไม่ได้รับการอนุมัติ', html)
        self.assertIn('BK999_TEST', html)
        self.assertIn('ยานพาหนะอยู่ระหว่างการซ่อมบำรุงประจำปี', html)

    def test_05_other_user_does_not_see_another_users_rejection_notification(self):
        """ตรวจสอบว่าผู้ใช้งานคนอื่น (เช่น นางสมศรี) จะต้องไม่เห็นการแจ้งเตือนคำขอที่ไม่อนุมัติของนายสมชาย"""
        # จำลองล็อกอินเป็นนางสมศรี
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
            sess['user_name'] = 'นางสมศรี สดใส'
            sess['email'] = 'somsri@farm.go.th'
            sess['role'] = 'user'

        res = self.client.get('/overview-calendar')
        self.assertEqual(res.status_code, 200)
        html = res.get_data(as_text=True)

        # นางสมศรีต้องไม่เห็นคำขอ BK999_TEST ของนายสมชาย
        self.assertNotIn('BK999_TEST', html)
        self.assertNotIn('ยานพาหนะอยู่ระหว่างการซ่อมบำรุงประจำปี', html)

if __name__ == '__main__':
    unittest.main()
