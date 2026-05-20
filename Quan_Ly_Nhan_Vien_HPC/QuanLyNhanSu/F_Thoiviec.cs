using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Views.Grid;
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
    public partial class F_Thoiviec : DevExpress.XtraEditors.XtraForm
    {
        public F_Thoiviec()
        {
            InitializeComponent();
        }
     
        private void LoadTrangThai()
        {
            cb_trangthai.Items.Clear();

            cb_trangthai.Items.Add("Đang làm việc");
            cb_trangthai.Items.Add("Đã nghỉ việc");
        }
        
        private void F_Thoiviec_Load(object sender, EventArgs e)
        {
            LoadTrangThai();
            loadData();
            gv_nhanvien.OptionsBehavior.Editable = true;
            gv_nhanvien.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None;
            GridView gridView = gv_nhanvien;
            // Định dạng căn giữa chữ trong Header
            gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            // Định dạng căn giữa chữ trong Row
            gridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            btnSua1.Click += Btn_sua_Click;
            
            if (gv_nhanvien.Columns["STT"] != null)
                gv_nhanvien.Columns["STT"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_nhanvien.Columns["MaNV"] != null)
                gv_nhanvien.Columns["MaNV"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_nhanvien.Columns["HoDem"] != null)
                gv_nhanvien.Columns["HoDem"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;
            if (gv_nhanvien.Columns["Ten"] != null)
                gv_nhanvien.Columns["Ten"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;
            if (gv_nhanvien.Columns["NgaySinh"] != null)
                gv_nhanvien.Columns["NgaySinh"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;
            if (gv_nhanvien.Columns["SDT"] != null)
                gv_nhanvien.Columns["SDT"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;


            if (gv_nhanvien.Columns["btnXoa1"] != null)
                gv_nhanvien.Columns["btnXoa1"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;

            if (gv_nhanvien.Columns["btnSua1"] != null)
                gv_nhanvien.Columns["btnSua1"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;

            // ===== Tự động độ rộng =====
            gv_nhanvien.OptionsView.ColumnAutoWidth = false;
            gv_nhanvien.BestFitColumns();

            // ===== Thanh cuộn =====
            gv_nhanvien.HorzScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;

            gv_nhanvien.VertScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            //gv_nhanvien.Columns["HoDem"].Width = 250;
        }

        private void loadData()
        {
            gc_nhanvien.DataSource = ConnectData.getdata("SELECT \r\n    nv.*,\r\n    hd.So_HopDong,\r\n    hd.NgayBD,\r\n    hd.LoaiHopDong,\r\n    hd.ChucVu,\r\n    hd.LuongCB,\r\n    hd.HeSoLuong,\r\n    pb.MaPhongBan,\r\n    pb.TenPhongBan\r\nFROM NHANVIEN nv\r\nLEFT JOIN HOPDONG hd \r\n    ON nv.id = hd.NhanVien_id\r\nLEFT JOIN PHONGBAN pb \r\n    ON hd.PhongBan_ID = pb.id\r\nWHERE nv.TrangThai = 'Đã nghỉ việc'\r\n   OR nv.DELETEO_BY IS NOT NULL;");

        }
        DangNhap Login = (DangNhap)Application.OpenForms["DangNhap"];
        private void Btn_sua_Click(object sender, EventArgs e)
        {
            if (gv_nhanvien.FocusedRowHandle < 0)
            {
                MessageBox.Show("Chọn nhân viên cần sửa");
                return;
            }

            string maNV = gv_nhanvien.GetFocusedRowCellValue("MaNV").ToString();

            string hoTen =
                gv_nhanvien.GetFocusedRowCellValue("HoDem") + " " +
                gv_nhanvien.GetFocusedRowCellValue("Ten");

            if (MessageBox.Show(
                "Bạn có chắc muốn khôi phục nhân viên:\n" + maNV + " - " + hoTen,
                "Thông báo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.No)
                return;

            try
            {
                ConnectData.taoketnoi();

                int id = Convert.ToInt32(
                    gv_nhanvien.GetFocusedRowCellValue("id"));

                string sql = @"
                        UPDATE NHANVIEN
                        SET
                            TrangThai = 'Đang làm việc',
                            DELETEO_BY = NULL,
                            DELETEO_DATE = NULL,
                            UPDATEO_BY = @UPDATEO_BY,
                            UPDATEO_DATE = @UPDATEO_DATE
                        WHERE id = @id";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@UPDATEO_BY",
                    Login.txt_taikhoan.Text);

                cmd.Parameters.AddWithValue(
                    "@UPDATEO_DATE",
                    DateTime.Now);

                cmd.Parameters.AddWithValue(
                    "@id",
                    id);

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Khôi phục thành công nhân viên:\n" +
                    maNV + " - " + hoTen,
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ConnectData.dongketnoi();
            }
            loadData();
        }

        private void btn_lammoi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loadData();
        }

        private void btn_dong_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            DialogResult dr;
            dr = XtraMessageBox.Show("Bạn có muốn đóng trang không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                this.Close();
            }
        }
    }
}