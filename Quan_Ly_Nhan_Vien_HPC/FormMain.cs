using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class FormMain : DevExpress.XtraEditors.XtraForm
    {
        public FormMain()
        {
            InitializeComponent();
        }
        public FormMain(string tk)
        {
            InitializeComponent();
            id_ho.Caption = tk;

        }
        void openForm(Type typeForm)
        {
            foreach (var frm in MdiChildren)
            {
                if (frm.GetType() == typeForm)
                {
                    frm.Activate();
                    return;
                }
            }
            Form f = (Form)Activator.CreateInstance(typeForm);
            f.MdiParent = this;
            f.Show();

        }


        private void listBox_sinhnhat_CustomizeItem(object sender, CustomizeTemplatedItemEventArgs e)
        {
           
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult rs = MessageBox.Show(
        "Bạn có muốn thoát không?",
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
                // Hiện lại form đăng nhập
                DangNhap login = new DangNhap();
                login.Show();
            }
        }

        private void ThongTin_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            loadSinhNhat();
            ls_sinhnhat.CustomizeItem += ls_sinhnhat_CustomizeItem;
            ribbonControl1.SelectedPage = NhanSu;
            AutoBackup.BackupDatabase();

        }

        private void ls_sinhnhat_CustomizeItem(object sender, CustomizeTemplatedItemEventArgs e)
        {
            if (e.TemplatedItem.Elements[1].Text.Substring(0, 2) == DateTime.Now.Day.ToString())
            {
                e.TemplatedItem.AppearanceItem.Normal.ForeColor = Color.Red;
            }
        }

        private void DoiMatKhau_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            F_DoiMatKhau f_Doimatkhau = new F_DoiMatKhau(id_ho.Caption);
            f_Doimatkhau.ShowDialog();
        }

        private void btn_thoat_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            DangNhap f_login = new DangNhap();

            DialogResult kq = MessageBox.Show("Bạn có muốn đăng xuất không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (kq == DialogResult.Yes)
            {
                //this.Hide();
                //f_DangNhap.ShowDialog();
                //f_DangNhap = null;
                this.Close();

            }
        }

        private void btn_phongban_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openForm(typeof(QuanLyPhongBan));
        }

        private void NhanVien_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openForm(typeof(NhanVien));
        }

        private void btn_BaoHiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openForm(typeof(F_BaoHiem));
        }

        private void HopDong_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openForm(typeof(F_HopDong1));
        }

        private void Nangluong_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openForm(typeof(F_GiaDinh));
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openForm(typeof(F_Doan_Dang));
        }

        private void ThoiViec_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openForm(typeof(F_Thoiviec));
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openForm(typeof(F_Tonghopthongtin));
        }
        void loadSinhNhat()
        {
            ls_sinhnhat.DataSource = ConnectData.getdata("SELECT \r\n    id, \r\n    CONCAT(HoDem, ' ', Ten) AS HoTen,\r\n    DATE_FORMAT(NgaySinh, '%d/%m/%Y') AS NgaySinh\r\nFROM NHANVIEN \r\nWHERE MONTH(NgaySinh) = MONTH(CURDATE())");
            ls_sinhnhat.DisplayMember = "HoTen";
            ls_sinhnhat.ValueMember = "id";
        }

        private void ribbonControl1_Click(object sender, EventArgs e)
        {

        }

        private void btn_backupdulieu_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                SaveFileDialog save =
                    new SaveFileDialog();

                save.Filter =
                    "SQL File|*.sql";

                save.Title =
                    "Backup dữ liệu";

                save.FileName =
                    "Backup_QuanLyNhanSu.sql";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                // ===== MYSQL INFO =====
                string host = ConnectData.server;

                string port = ConnectData.port;

                string database = ConnectData.database;

                string user = ConnectData.user;

                string password = ConnectData.password;
                    

                // ===== LINUX MYSQLDUMP =====
                string mysqldump = @"C:\xampp\mysql\bin\mysqldump.exe";

                // ===== ARGUMENT =====
                string arguments =
                    $"-h {host} -P {port} -u {user} -p{password} {database} --result-file=\"{save.FileName}\"";

                ProcessStartInfo psi =
                    new ProcessStartInfo();

                psi.FileName =
                    mysqldump;

                psi.Arguments =
                    arguments;

                psi.RedirectStandardError =
                    true;

                psi.RedirectStandardOutput =
                    true;

                psi.UseShellExecute =
                    false;

                psi.CreateNoWindow =
                    true;

                Process process =
                    new Process();

                process.StartInfo =
                    psi;

                process.Start();

                process.WaitForExit();

                MessageBox.Show(
                    "Backup dữ liệu thành công",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_log_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                string folder =
                    Application.StartupPath + @"\Logs";

                if (!Directory.Exists(folder))
                {
                    MessageBox.Show(
                        "Chưa có log");

                    return;
                }

                System.Diagnostics.Process.Start(folder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}