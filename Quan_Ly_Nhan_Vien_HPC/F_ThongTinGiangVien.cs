using DevExpress.XtraEditors;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class F_ThongTinGiangVien : DevExpress.XtraEditors.XtraForm
    {
        public F_ThongTinGiangVien()
        {
            InitializeComponent();
        }
        string matk;
        public F_ThongTinGiangVien(string manv)
        {
            //InitializeComponent();
            ////matk = manv;
            //txt_hodem.Text = manv;

        }
        NhanVien f_GiangVien = (NhanVien)Application.OpenForms["NhanVien"];
        DangNhap Login = (DangNhap)Application.OpenForms["DangNhap"];
        string imagePath = "";
        private void btn_chonHA_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();

            open.Filter =
                "Image Files|*.jpg;*.jpeg;*.png;*.bmp";

            if (open.ShowDialog() == DialogResult.OK)
            {
                // Thư mục Images cùng phần mềm
                string folder =
                    Application.StartupPath + @"\Images\";

                // Tạo thư mục nếu chưa có
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                // Tên file
                string fileName =
                    Path.GetFileName(open.FileName);

                // Đường dẫn mới
                imagePath = folder + fileName;

                // Copy ảnh
                File.Copy(open.FileName, imagePath, true);

                // Hiển thị ảnh
                pcb_hinhanh.Image =
                    Image.FromFile(imagePath);
            }
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void F_ThongTinGiangVien_Load(object sender, EventArgs e)
        {
            //txt_hodem.Text = Login.txt_taikhoan.Text;
            cb_gioitimh.Text = "Nam";
            txt_hodem.KeyDown += txt_hodem_KeyDown;
            txt_tennv.KeyDown += txt_tennv_KeyDown;
            date_ngaysinh.KeyDown += date_ngaysinh_KeyDown;
            cb_gioitimh.KeyDown += cb_gioitimh_KeyDown;
            txt_quoctich.KeyDown += txt_quoctich_KeyDown;
            txt_quequan.KeyDown += txt_quequan_KeyDown;
            txt_cccd.KeyDown += txt_cccd_KeyDown;
            txt_noicap.KeyDown += txt_noicap_KeyDown;
            date_ngaycap.KeyDown += date_ngaycap_KeyDown;
            txt_dantoc.KeyDown += txt_dantoc_KeyDown;
            txt_tongiao.KeyDown += txt_tongiao_KeyDown;
            txt_sdt.KeyDown += txt_sdt_KeyDown;
            txt_dcThuongChu.KeyDown += txt_dcThuongChu_KeyDown;
            txt_trinhdochuyenmon.KeyDown += txt_trinhdochuyenmon_KeyDown;
            txt_GDPT.KeyDown += txt_GDPT_KeyDown;
            txt_HocVi.KeyDown += txt_HocVi_KeyDown;
            txt_chungchikhac.KeyDown += txt_chungchikhac_KeyDown;
            txt_ghichu.KeyDown += txt_ghichu_KeyDown;
            txt_email.KeyDown += txt_email_KeyDown;
            txt_trinhdochuyenmon.KeyDown += txt_trinhdochuyenmon_KeyDown;
            txt_masothue.KeyDown += txt_masothue_KeyDown;

        }

        private void txt_masothue_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_email.Focus();
            }
        }

        private void txt_email_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_GDPT.Focus();
            }
        }

        private void txt_ghichu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_luu.PerformClick();
            }
        }

        private void txt_chungchikhac_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_ghichu.Focus();
            }
        }

        private void txt_HocVi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_chungchikhac.Focus();
            }
        }

        private void txt_GDPT_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_HocVi.Focus();
            }
        }

        private void txt_trinhdochuyenmon_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_dantoc.Focus();
            }
        }

        private void txt_dcThuongChu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_trinhdochuyenmon.Focus();
            }
        }

        private void txt_sdt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_dcThuongChu.Focus();
            }
        }

        private void txt_noicap_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                date_ngaycap.Focus();
            }
        }

        private void txt_cccd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_noicap.Focus();
            }
        }

        private void date_ngaycap_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_sdt.Focus();
            }
        }

        private void txt_tongiao_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_masothue.Focus();
            }
        }

        private void txt_dantoc_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_tongiao.Focus();
            }
        }

        private void txt_quequan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_cccd.Focus();
            }
        }

        private void txt_quoctich_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_quequan.Focus();
            }
        }

        private void cb_gioitimh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_quoctich.Focus();
            }
        }

        private void date_ngaysinh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                cb_gioitimh.Focus();
            }
        }

        private void txt_tennv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                date_ngaysinh.Focus();
            }
        }

        private void txt_hodem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_tennv.Focus();
            }
        }

        private void btn_luu_Click(object sender, EventArgs e)
        {
            // ===== Kiểm tra rỗng =====
            if (txt_hodem.Text.Trim() == "")
            {
                MessageBox.Show("Nhập họ đệm");
                txt_hodem.Focus();
                return;
            }

            if (txt_tennv.Text.Trim() == "")
            {
                MessageBox.Show("Nhập tên");
                txt_tennv.Focus();
                return;
            }

            

            if (cb_gioitimh.Text.Trim() == "")
            {
                MessageBox.Show("Chọn giới tính");
                cb_gioitimh.Focus();
                return;
            }

            if (txt_dantoc.Text.Trim() == "")
            {
                MessageBox.Show("Nhập dân tộc");
                txt_dantoc.Focus();
                return;
            }

            if (txt_quoctich.Text.Trim() == "")
            {
                MessageBox.Show("Nhập quốc tịch");
                txt_quoctich.Focus();
                return;
            }

            if (txt_tongiao.Text.Trim() == "")
            {
                MessageBox.Show("Nhập tôn giáo");
                txt_tongiao.Focus();
                return;
            }

            if (txt_quequan.Text.Trim() == "")
            {
                MessageBox.Show("Nhập quê quán");
                txt_quequan.Focus();
                return;
            }

            if (txt_dcThuongChu.Text.Trim() == "")
            {
                MessageBox.Show("Nhập địa chỉ");
                txt_dcThuongChu.Focus();
                return;
            }

            if (txt_cccd.Text.Trim() == "")
            {
                MessageBox.Show("Nhập CCCD");
                txt_cccd.Focus();
                return;
            }

            if (txt_noicap.Text.Trim() == "")
            {
                MessageBox.Show("Nhập nơi cấp CCCD");
                txt_noicap.Focus();
                return;
            }

            if (txt_sdt.Text.Trim() == "")
            {
                MessageBox.Show("Nhập số điện thoại");
                txt_sdt.Focus();
                return;
            }

            if (txt_email.Text.Trim() == "")
            {
                MessageBox.Show("Nhập email");
                txt_email.Focus();
                return;
            }

            if (txt_GDPT.Text.Trim() == "")
            {
                MessageBox.Show("Nhập trình độ giáo dục phổ thông");
                txt_GDPT.Focus();
                return;
            }

            if (txt_trinhdochuyenmon.Text.Trim() == "")
            {
                MessageBox.Show("Nhập trình độ chuyên môn");
                txt_trinhdochuyenmon.Focus();
                return;
            }

            if (txt_HocVi.Text.Trim() == "")
            {
                MessageBox.Show("Nhập học hàm học vị");
                txt_HocVi.Focus();
                return;
            }

            try
            {
                ConnectData.taoketnoi();

                // ===== Tạo mã nhân viên tự tăng =====
                string maNV = "26000000";

                string sqlMa =
                    "SELECT MAX(CAST(MaNV AS UNSIGNED)) FROM NHANVIEN";

                MySqlCommand cmdMa =
                    new MySqlCommand(sqlMa, ConnectData.conn);

                object result = cmdMa.ExecuteScalar();

                if (result != DBNull.Value && result != null)
                {
                    maNV =
                        (Convert.ToInt64(result) + 1).ToString();
                }

                // ===== Check CCCD =====
                string sqlCCCD =
                    "SELECT COUNT(*) FROM NHANVIEN WHERE CCCD=@CCCD";

                MySqlCommand cmdCCCD =
                    new MySqlCommand(sqlCCCD, ConnectData.conn);

                cmdCCCD.Parameters.AddWithValue(
                    "@CCCD",
                    txt_cccd.Text.Trim());

                int checkCCCD =
                    Convert.ToInt32(cmdCCCD.ExecuteScalar());

                if (checkCCCD > 0)
                {
                    MessageBox.Show("CCCD đã tồn tại");
                    txt_cccd.Focus();
                    return;
                }

                // ===== Check Email =====
                string sqlEmail =
                    "SELECT COUNT(*) FROM NHANVIEN WHERE Email=@Email";

                MySqlCommand cmdEmail =
                    new MySqlCommand(sqlEmail, ConnectData.conn);

                cmdEmail.Parameters.AddWithValue(
                    "@Email",
                    txt_email.Text.Trim());

                int checkEmail =
                    Convert.ToInt32(cmdEmail.ExecuteScalar());

                if (checkEmail > 0)
                {
                    MessageBox.Show("Email đã tồn tại");
                    txt_email.Focus();
                    return;
                }

                // ===== Check SĐT =====
                string sqlSDT =
                    "SELECT COUNT(*) FROM NHANVIEN WHERE SDT=@SDT";

                MySqlCommand cmdSDT =
                    new MySqlCommand(sqlSDT, ConnectData.conn);

                cmdSDT.Parameters.AddWithValue(
                    "@SDT",
                    txt_sdt.Text.Trim());

                int checkSDT =
                    Convert.ToInt32(cmdSDT.ExecuteScalar());

                if (checkSDT > 0)
                {
                    MessageBox.Show("Số điện thoại đã tồn tại");
                    txt_sdt.Focus();
                    return;
                }

                // ===== Insert =====
                string sql = @"INSERT INTO NHANVIEN
        (
            MaNV,
            HoDem,
            Ten,
            Password,
            NgaySinh,
            GioiTinh,
            DanToc,
            QuocTich,
            TonGiao,
            QueQuan,
            DiaChiThuongChu,
            CCCD,
            NoiCapCCCD,
            NgayCapCCCD,
            SDT,
            Email,
            TrinhDoGiaoDucPhoThong,
            TrinhDoChuyenMon,
            HocHamHocVi,
            ChungChiKhac,
            HinhAnh,
            MaSoThue,
            TrangThai,
            GhiChu,
            CREATEO_BY,
            CREATEO_DATE
        )
        VALUES
        (
            @MaNV,
            @HoDem,
            @Ten,
            @Password,
            @NgaySinh,
            @GioiTinh,
            @DanToc,
            @QuocTich,
            @TonGiao,
            @QueQuan,
            @DiaChiThuongChu,
            @CCCD,
            @NoiCapCCCD,
            @NgayCapCCCD,
            @SDT,
            @Email,
            @TrinhDoGiaoDucPhoThong,
            @TrinhDoChuyenMon,
            @HocHamHocVi,
            @ChungChiKhac,
            @HinhAnh,
            @MaSoThue,
            @TrangThai,
            @GhiChu,
            @CREATEO_BY,
            @CREATEO_DATE
        )";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue("@MaNV", maNV);

                cmd.Parameters.AddWithValue("@HoDem",
                    System.Globalization.CultureInfo
                    .CurrentCulture.TextInfo
                    .ToTitleCase(txt_hodem.Text.Trim().ToLower()));

                cmd.Parameters.AddWithValue("@Ten",
                    System.Globalization.CultureInfo
                    .CurrentCulture.TextInfo
                    .ToTitleCase(txt_tennv.Text.Trim().ToLower()));

                cmd.Parameters.AddWithValue("@Password", "12345678");

                cmd.Parameters.AddWithValue("@NgaySinh", date_ngaysinh.Value);

                cmd.Parameters.AddWithValue("@GioiTinh", cb_gioitimh.Text);

                cmd.Parameters.AddWithValue("@DanToc", txt_dantoc.Text.Trim());

                cmd.Parameters.AddWithValue("@QuocTich", txt_quoctich.Text.Trim());

                cmd.Parameters.AddWithValue("@TonGiao", txt_tongiao.Text.Trim());

                cmd.Parameters.AddWithValue("@QueQuan", txt_quequan.Text.Trim());

                cmd.Parameters.AddWithValue("@DiaChiThuongChu", txt_dcThuongChu.Text.Trim());

                cmd.Parameters.AddWithValue("@CCCD", txt_cccd.Text.Trim());

                cmd.Parameters.AddWithValue("@NoiCapCCCD", txt_noicap.Text.Trim());

                cmd.Parameters.AddWithValue("@NgayCapCCCD", date_ngaycap.Value);

                cmd.Parameters.AddWithValue("@SDT", txt_sdt.Text.Trim());

                cmd.Parameters.AddWithValue("@Email", txt_email.Text.Trim());

                cmd.Parameters.AddWithValue("@TrinhDoGiaoDucPhoThong", txt_GDPT.Text.Trim());

                cmd.Parameters.AddWithValue("@TrinhDoChuyenMon", txt_trinhdochuyenmon.Text.Trim());

                cmd.Parameters.AddWithValue("@HocHamHocVi", txt_HocVi.Text.Trim());

                cmd.Parameters.AddWithValue("@ChungChiKhac", txt_chungchikhac.Text.Trim());

                cmd.Parameters.AddWithValue("@HinhAnh", imagePath);

                // Mã số thuế = CCCD
                cmd.Parameters.AddWithValue("@MaSoThue", txt_masothue.Text.Trim());

                cmd.Parameters.AddWithValue("@TrangThai", "Đang làm việc");

                cmd.Parameters.AddWithValue("@GhiChu", txt_ghichu.Text.Trim());

                cmd.Parameters.AddWithValue("@CREATEO_BY",Login.txt_taikhoan.Text);

                cmd.Parameters.AddWithValue("@CREATEO_DATE",DateTime.Now);

                cmd.ExecuteNonQuery();

                MessageBox.Show("Thêm nhân viên thành công\nMã NV: " + maNV);
                clearForm();
                
            }
            catch (MySqlException ex)
            {
                if (ex.Number == 1062)
                {
                    MessageBox.Show("Dữ liệu đã tồn tại");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ConnectData.dongketnoi();
            }
            f_GiangVien.loadData();
        }
        void clearForm()
        {
            txt_hodem.Clear();
            txt_tennv.Clear();
            date_ngaysinh.Value = DateTime.Now;
            cb_gioitimh.Text = "Nam";
            txt_quoctich.Clear();
            txt_quequan.Clear();
            txt_cccd.Clear();
            txt_noicap.Clear();
            date_ngaycap.Value = DateTime.Now;
            txt_sdt.Clear();
            txt_dcThuongChu.Clear();
            txt_trinhdochuyenmon.Clear();
            txt_GDPT.Clear();
            txt_HocVi.Clear();
            txt_chungchikhac.Clear();
            txt_ghichu.Clear();
            txt_email.Clear();
            pcb_hinhanh.Image = null;
        }
        private void txt_email_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_cccd_TextChanged(object sender, EventArgs e)
        {
            txt_masothue.Text = txt_cccd.Text ;
        }

        private void txt_tongiao_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void txt_dantoc_TextChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void btn_huy1_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("Bạn có muốn đóng trang không?", "Thông báo", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                clearForm();
                this.Close();
            }
        }
    }
}