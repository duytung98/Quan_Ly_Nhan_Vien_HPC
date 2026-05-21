using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class DangNhap : Form
    {
        public DangNhap()
        {
            InitializeComponent();
        }
        public string temp;
        private void DangNhap_Load(object sender, EventArgs e)
        {
            txt_taikhoan.KeyDown += txt_taikhoan_KeyDown;
            txt__matkhau.KeyDown += txt__matkhau_KeyDown;
        }

        private void txt__matkhau_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_dangnhap.PerformClick();
            }
        }

        private void txt_taikhoan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ
                
                txt__matkhau.Focus();
            }
        }

        public void login()
        {
            try
            {
                // Kiểm tra rỗng
                if (txt_taikhoan.Text.Trim() == "")
                {
                    MessageBox.Show("Vui lòng nhập tài khoản");
                    txt_taikhoan.Focus();
                    return;
                }

                if (txt__matkhau.Text.Trim() == "")
                {
                    MessageBox.Show("Vui lòng nhập mật khẩu");
                    txt__matkhau.Focus();
                    return;
                }

                // Câu lệnh SQL
                string sql = @"SELECT COUNT(*) 
                       FROM Admin 
                       WHERE Username = @Username 
                       AND Password = @Password";

                // Parameter
                MySqlParameter[] param =
                {
            new MySqlParameter("@Username", txt_taikhoan.Text.Trim()),
            new MySqlParameter("@Password", txt__matkhau.Text.Trim())
        };

                // Thực thi
                object result = ConnectData.execScalar(sql, param);

                int count = Convert.ToInt32(result);

                if (count > 0)
                {
                    // Mở form chính
                    FormMain fm = new FormMain(txt_taikhoan.Text);
                    fm.Show();

                    // Ẩn form login
                    this.Hide();
                    temp = txt_taikhoan.Text;
                    LogSystem.CurrentUser = txt_taikhoan.Text;
                }
                else
                {
                    MessageBox.Show("Sai tài khoản hoặc mật khẩu",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    txt__matkhau.Clear();
                    txt__matkhau.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi đăng nhập: " + ex.Message);
            }
        }
        private void btn_dangnhap_Click(object sender, EventArgs e)
        {
            login();
        }

        private void ckc_hienthimk_CheckedChanged(object sender, EventArgs e)
        {
            if (ckc_hienthimk.Checked)
            {
                txt__matkhau.UseSystemPasswordChar = true; // hiện mật khẩu
            }
            else
            {
                txt__matkhau.UseSystemPasswordChar = false; // ẩn mật khẩu
            }
        }

        private void btn_thoat_Click(object sender, EventArgs e)
        {
          
        }

        private void DangNhap_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult rs = MessageBox.Show(
       "Bạn có muốn thoát chương trình không?",
       "Thông báo",
       MessageBoxButtons.YesNo,
       MessageBoxIcon.Question);

            if (rs == DialogResult.No)
            {
                // Hủy thoát
                e.Cancel = true;
            }
            else
            {
                // Thoát toàn bộ chương trình
                Application.Exit();
            }
        }
    }
}
