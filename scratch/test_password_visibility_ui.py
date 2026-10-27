import unittest
from app import app

class TestPasswordVisibilityUI(unittest.TestCase):
    def setUp(self):
        self.app = app
        self.app.config['TESTING'] = True
        self.client = self.app.test_client()

    def test_login_page_password_wrappers(self):
        """ตรวจสอบว่าหน้า Login มีการใช้ .password-field-wrapper และปุ่มตาอยู่ด้านในช่องรหัสผ่าน"""
        res = self.client.get('/login')
        self.assertEqual(res.status_code, 200)
        html = res.get_data(as_text=True)

        # ต้องมี .password-field-wrapper
        self.assertIn('password-field-wrapper', html)

        # ต้องมี input สำหรับรหัสผ่านแต่ละจุดครบถ้วน
        self.assertIn('id="login_password"', html)
        self.assertIn('id="reg_password"', html)
        self.assertIn('id="reg_confirm_password"', html)
        self.assertIn('id="forgot_new_password"', html)
        self.assertIn('id="forgot_confirm_password"', html)

        # ตรวจสอบว่าช่อง password แต่ละช่องไม่มี style ตัดขอบมนทิ้ง
        self.assertNotIn('id="login_password"\n                                           placeholder="กรอกรหัสผ่านอย่างน้อย 8 ตัวอักษร"\n                                           minlength="8"\n                                           required\n                                           oninput="updateCharCount(\'login_password\', \'loginPwdCount\', 8)"\n                                           style="border-top-right-radius: 0;', html)

        # ปุ่มเปิดตา toggle-password-btn ต้องอยู่ด้านใน
        self.assertIn('class="toggle-password-btn"', html)
        self.assertIn('class="eye-icon">👁️</span>', html)

    def test_admin_login_page_password_wrapper(self):
        """ตรวจสอบว่าหน้า Admin Login มีการใช้ .password-field-wrapper และปุ่มตาอยู่ด้านในช่องรหัสผ่าน"""
        res = self.client.get('/approve/login')
        self.assertEqual(res.status_code, 200)
        html = res.get_data(as_text=True)

        self.assertIn('password-field-wrapper', html)
        self.assertIn('id="admin_password"', html)
        self.assertIn('class="toggle-password-btn"', html)
        self.assertIn('class="eye-icon">👁️</span>', html)

    def test_farm_theme_css_password_rules(self):
        """ตรวจสอบว่าไฟล์ CSS farm-theme.css มีการกำหนดสไตล์ช่องรหัสผ่านและปุ่มตาแบบฝังในช่องครบถ้วน"""
        with open('static/css/farm-theme.css', 'r', encoding='utf-8') as f:
            css = f.read()

        self.assertIn('.password-field-wrapper', css)
        self.assertIn('padding-right: 46px', css)
        self.assertIn('.toggle-password-btn', css)
        self.assertIn('transform: translateY(-50%)', css)
        self.assertIn('background: transparent', css)

if __name__ == '__main__':
    unittest.main()
