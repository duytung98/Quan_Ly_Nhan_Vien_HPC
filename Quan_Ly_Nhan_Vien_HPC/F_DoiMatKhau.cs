using DevExpress.XtraEditors;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class F_DoiMatKhau : DevExpress.XtraEditors.XtraForm
    {
        public F_DoiMatKhau()
        {
            InitializeComponent();
        }
        string matk;
        public F_DoiMatKhau(string manv)
        {
            InitializeComponent();
            matk = manv;
        }
        private void btn_mat2_Click(object sender, EventArgs e)
        {
            if (txt_mkmoi.PasswordChar == '*')
            {
                txt_mkmoi.PasswordChar = '\0';
            }
            else
            {
                txt_mkmoi.PasswordChar = '*';
            }
        }

        private void btn_mat3_Click(object sender, EventArgs e)
        {
            if (txt_xnmk.PasswordChar == '*')
            {
                txt_xnmk.PasswordChar = '\0';
            }
            else
            {
                txt_xnmk.PasswordChar = '*';
            }
        }
        ConnectData c = new ConnectData();
        string email;
        Random Random = new Random();
        int otp;
        private void btn_guiotp_Click(object sender, EventArgs e)
        {
            ConnectData.taoketnoi();
            MySqlCommand cmd = new MySqlCommand("SELECT * FROM Admin where Username = '" + matk + "';", ConnectData.conn);
            MySqlDataReader reader = cmd.ExecuteReader();
            try
            {
                while (reader.Read())
                {
                    email = reader["Email"].ToString();
                    try
                    {
                        otp = Random.Next(100000, 1000000);
                        var fromAddress = new MailAddress("duytung1998vp@gmail.com");
                        var toAddress = new MailAddress(email);
                        const string frompass = "xmev jevm gbya posn";
                        const string subject = "OTP code ";
                        string body = otp.ToString();
                        var smtp = new SmtpClient
                        {
                            Host = "smtp.gmail.com",
                            Port = 587,
                            EnableSsl = true,
                            DeliveryMethod = SmtpDeliveryMethod.Network,
                            UseDefaultCredentials = false,
                            Credentials = new NetworkCredential(fromAddress.Address, frompass),
                            Timeout = 200000

                        };
                        using (var messger = new MailMessage(fromAddress, toAddress)
                        {
                            Subject = subject,
                            Body = body
                        })
                        {
                            smtp.Send(messger);
                        }
                        MessageBox.Show("Mã OTP đã được gửi qua email đăng ký", "Thông báo");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            finally
            {
                reader.Close();
            }
            ConnectData.dongketnoi();

        }

        private void btb_dongy_Click(object sender, EventArgs e)
        {
            if (txtOTP.Text == "" || txt_mkmoi.Text == "" || txt_xnmk.Text == "")
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (txt_mkmoi.Text.Length < 8)
            {
                MessageBox.Show("Mật khẩu phải từ 8 ký tự trở lên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (txt_xnmk.Text != txt_mkmoi.Text)
            {
                MessageBox.Show(",Xác nhận mật khẩu sai, Vui lòng nhập lại mật khẩu!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (otp.ToString().Equals(txtOTP.Text))
            {
                ConnectData.execQuery("UPDATE `Admin` SET `Password`='" + txt_xnmk.Text + "' WHERE `Username`= '" + matk + "';");
                MessageBox.Show("Đổi mật khẩu thành công!!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtOTP.Text = "";
                txt_mkmoi.Text = "";
                txt_xnmk.Text = "";
                this.Close();
            }
            else
            {
                MessageBox.Show("Vui lòng kiểm tra lại mã OTP", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_thoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void F_DoiMatKhau_Load(object sender, EventArgs e)
        {

        }
    }
}