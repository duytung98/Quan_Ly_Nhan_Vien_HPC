using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
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
            if (e.TemplatedItem.Elements[1].Text.Substring(0,2) == DateTime.Now.Day.ToString())
            {
                e.TemplatedItem.AppearanceItem.Normal.ForeColor = Color.Red;
            }
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
    }
}