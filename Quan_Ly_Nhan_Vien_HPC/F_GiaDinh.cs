using DevExpress.XtraEditors;
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
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class F_GiaDinh : DevExpress.XtraEditors.XtraForm
    {
        public F_GiaDinh()
        {
            InitializeComponent();
        }

        private void F_GiaDinh_Load(object sender, EventArgs e)
        {
            load_thongtingiadinh();
            lku_nhanvien.Focus();
            
            Loadnhanvien();
            gv_giadinh.OptionsBehavior.Editable = true;
            gv_giadinh.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None;
            GridView gridView = gv_giadinh;
            // Định dạng căn giữa chữ trong Header
            gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            // Định dạng căn giữa chữ trong Row
            gridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            btn_sua.Click += Btn_sua_Click;
            btn_xoa.Click += Btn_xoa_Click;
            //lku_nhanvien.KeyDown += lku_nhanvien_KeyDown;
            //txt_luongcb.KeyDown += txt_luongcb_KeyDown;
            //txt_ghichu.KeyDown += txt_ghichu_KeyDown;
            //txt_hesoluong.KeyDown += txt_hesoluong_KeyDown;
            //cb_loaihopdong.KeyDown += cb_loaihopdong_KeyDown;
            //date_NgayBD.KeyDown += date_NgayBD_KeyDown;
            //txt_chucvu.KeyDown += txt_chucvu_KeyDown;
            //lbl_phongban.KeyDown += lbl_phongban_KeyDown;
            if (gv_giadinh.Columns["STT"] != null)
                gv_giadinh.Columns["STT"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_giadinh.Columns["MaNV"] != null)
                gv_giadinh.Columns["MaNV"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_giadinh.Columns["HoTen"] != null)
                gv_giadinh.Columns["HoTen"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_giadinh.Columns["btn_xoa"] != null)
                gv_giadinh.Columns["btn_xoa"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;

            if (gv_giadinh.Columns["btn_sua"] != null)
                gv_giadinh.Columns["btn_sua"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;



            // ===== Tự động độ rộng =====
            gv_giadinh.OptionsView.ColumnAutoWidth = false;
            gv_giadinh.BestFitColumns();

            // ===== Thanh cuộn =====
            gv_giadinh.HorzScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;

            gv_giadinh.VertScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            gv_giadinh.Columns["HoTen"].Width = 200;
            cb_hocvan1.SelectedIndex = -1;
            cb_hocvan2.SelectedIndex = -1;
            cb_hocvan3.SelectedIndex = -1;

            txt_ngaysinhvo.Properties.Mask.MaskType =DevExpress.XtraEditors.Mask.MaskType.DateTime;
            txt_ngaysinhvo.Properties.Mask.EditMask ="dd/MM/yyyy";
            txt_ngaysinhvo.Properties.Mask.UseMaskAsDisplayFormat =true;

            date_con1.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTime;
            date_con1.Properties.Mask.EditMask = "dd/MM/yyyy";
            date_con1.Properties.Mask.UseMaskAsDisplayFormat = true;

            date_con2.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTime;
            date_con2.Properties.Mask.EditMask = "dd/MM/yyyy";
            date_con2.Properties.Mask.UseMaskAsDisplayFormat = true;

            date_con3.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTime;
            date_con3.Properties.Mask.EditMask = "dd/MM/yyyy";
            date_con3.Properties.Mask.UseMaskAsDisplayFormat = true;
        }

        private void Btn_xoa_Click(object sender, EventArgs e)
        {
            try
            {
                // ===== Check chọn dòng =====
                if (gv_giadinh.FocusedRowHandle < 0)
                {
                    MessageBox.Show(
                        "Vui lòng chọn dữ liệu cần xóa",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // ===== Lấy ID =====
                int id =
                    Convert.ToInt32(
                        gv_giadinh.GetFocusedRowCellValue("Id"));

                // ===== Lấy họ tên nhân viên =====
                string hotennhanvien =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("HoTen"));

                // ===== Hỏi xác nhận =====
                DialogResult rs =
                    MessageBox.Show(
                        "Bạn có chắc chắn muốn xóa thông tin gia đình của nhân viên: "
                        + hotennhanvien + " ?",
                        "Thông báo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                if (rs != DialogResult.Yes)
                    return;

                ConnectData.taoketnoi();

                // ===== SQL DELETE =====
                string sql =
                    "DELETE FROM THONGTINGIADINH WHERE Id=@Id";

                MySqlCommand cmd =
                    new MySqlCommand(
                        sql,
                        ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@Id",
                    id);

                int kq =
                    cmd.ExecuteNonQuery();

                if (kq > 0)
                {
                    MessageBox.Show(
                        "Xóa thông tin gia đình thành công",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "Xóa thất bại",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                ConnectData.dongketnoi();
            }

            // ===== Reload dữ liệu =====
            load_thongtingiadinh();
        }

        private void Btn_sua_Click(object sender, EventArgs e)
        {
            try
            {
                if (gv_giadinh.FocusedRowHandle < 0)
                    return;

                DialogResult rs =
                    MessageBox.Show(
                        "Bạn có chắc chắn muốn sửa thông tin gia đình này không?",
                        "Thông báo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                if (rs != DialogResult.Yes)
                    return;

                // ===== ID =====
                int id =
                    Convert.ToInt32(
                        gv_giadinh.GetFocusedRowCellValue("Id"));

                // ===== Bố =====
                string hoTenBo =
                    VietHoaChuCaiDau(
                        Convert.ToString(
                            gv_giadinh.GetFocusedRowCellValue("HoTenBo")));

                DateTime ngaySinhBo =
                    Convert.ToDateTime(
                        gv_giadinh.GetFocusedRowCellValue("NgaySinhBo"));

                string ngheNghiepBo =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("NgheNghiepBo"));

                string sdtBo =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("SDTBo"));

                // ===== Mẹ =====
                string hoTenMe =
                    VietHoaChuCaiDau(
                        Convert.ToString(
                            gv_giadinh.GetFocusedRowCellValue("HoTenMe")));

                DateTime ngaySinhMe =
                    Convert.ToDateTime(
                        gv_giadinh.GetFocusedRowCellValue("NgaySinhMe"));

                string ngheNghiepMe =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("NgheNghiepMe"));

                string sdtMe =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("SDTMe"));

                // ===== Vợ / Chồng =====
                string hoTenVo =
                    VietHoaChuCaiDau(
                        Convert.ToString(
                            gv_giadinh.GetFocusedRowCellValue("HoTenVo_Chong")));

                object ngayVoObj =
                    gv_giadinh.GetFocusedRowCellValue("NgaySinhVo_Chong");

                DateTime? ngaySinhVo = null;

                if (ngayVoObj != null
                    && ngayVoObj != DBNull.Value
                    && !string.IsNullOrWhiteSpace(
                        ngayVoObj.ToString()))
                {
                    ngaySinhVo =
                        Convert.ToDateTime(ngayVoObj);
                }

                string sdtVo =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("SDTVo_Chong"));

                string ngheNghiepVo =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("NgheNghiepVo_Chong"));

                // ===== Con 1 =====
                string hoTenCon1 =
                    VietHoaChuCaiDau(
                        Convert.ToString(
                            gv_giadinh.GetFocusedRowCellValue("HoTenCon1")));

                object ngayCon1Obj =
                    gv_giadinh.GetFocusedRowCellValue("NgaySinhCon1");

                DateTime? ngaySinhCon1 = null;

                if (ngayCon1Obj != null
                    && ngayCon1Obj != DBNull.Value
                    && !string.IsNullOrWhiteSpace(
                        ngayCon1Obj.ToString()))
                {
                    ngaySinhCon1 =
                        Convert.ToDateTime(ngayCon1Obj);
                }

                string hocVanCon1 =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("HocVanCon1"));

                // ===== Con 2 =====
                string hoTenCon2 =
                    VietHoaChuCaiDau(
                        Convert.ToString(
                            gv_giadinh.GetFocusedRowCellValue("HoTenCon2")));

                object ngayCon2Obj =
                    gv_giadinh.GetFocusedRowCellValue("NgaySinhCon2");

                DateTime? ngaySinhCon2 = null;

                if (ngayCon2Obj != null
                    && ngayCon2Obj != DBNull.Value
                    && !string.IsNullOrWhiteSpace(
                        ngayCon2Obj.ToString()))
                {
                    ngaySinhCon2 =
                        Convert.ToDateTime(ngayCon2Obj);
                }

                string hocVanCon2 =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("HocVanCon2"));

                // ===== Con 3 =====
                string hoTenCon3 =
                    VietHoaChuCaiDau(
                        Convert.ToString(
                            gv_giadinh.GetFocusedRowCellValue("HoTenCon3")));

                object ngayCon3Obj =
                    gv_giadinh.GetFocusedRowCellValue("NgaySinhCon3");

                DateTime? ngaySinhCon3 = null;

                if (ngayCon3Obj != null
                    && ngayCon3Obj != DBNull.Value
                    && !string.IsNullOrWhiteSpace(
                        ngayCon3Obj.ToString()))
                {
                    ngaySinhCon3 =
                        Convert.ToDateTime(ngayCon3Obj);
                }

                string hocVanCon3 =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("HocVanCon3"));

                // ===== Ghi chú =====
                string ghiChu =
                    Convert.ToString(
                        gv_giadinh.GetFocusedRowCellValue("GhiChu"));

                // ===== Check SĐT =====
                if (!string.IsNullOrWhiteSpace(sdtBo)
                    && !sdtBo.All(char.IsDigit))
                {
                    MessageBox.Show(
                        "SĐT bố chỉ được nhập số");

                    return;
                }

                if (!string.IsNullOrWhiteSpace(sdtMe)
                    && !sdtMe.All(char.IsDigit))
                {
                    MessageBox.Show(
                        "SĐT mẹ chỉ được nhập số");

                    return;
                }

                if (!string.IsNullOrWhiteSpace(sdtVo)
                    && !sdtVo.All(char.IsDigit))
                {
                    MessageBox.Show(
                        "SĐT vợ/chồng chỉ được nhập số");

                    return;
                }

                ConnectData.taoketnoi();

                string sql = @"
UPDATE THONGTINGIADINH
SET
    HoTenBo=@HoTenBo,
    NgaySinhBo=@NgaySinhBo,
    NgheNghiepBo=@NgheNghiepBo,
    SDTBo=@SDTBo,

    HoTenMe=@HoTenMe,
    NgaySinhMe=@NgaySinhMe,
    NgheNghiepMe=@NgheNghiepMe,
    SDTMe=@SDTMe,

    HoTenVo_Chong=@HoTenVo_Chong,
    SDTVo_Chong=@SDTVo_Chong,
    NgaySinhVo_Chong=@NgaySinhVo_Chong,
    NgheNghiepVo_Chong=@NgheNghiepVo_Chong,

    HoTenCon1=@HoTenCon1,
    NgaySinhCon1=@NgaySinhCon1,
    HocVanCon1=@HocVanCon1,

    HoTenCon2=@HoTenCon2,
    NgaySinhCon2=@NgaySinhCon2,
    HocVanCon2=@HocVanCon2,

    HoTenCon3=@HoTenCon3,
    NgaySinhCon3=@NgaySinhCon3,
    HocVanCon3=@HocVanCon3,

    GhiChu=@GhiChu
WHERE Id=@Id";

                MySqlCommand cmd =
                    new MySqlCommand(
                        sql,
                        ConnectData.conn);

                cmd.Parameters.AddWithValue("@Id", id);

                cmd.Parameters.AddWithValue("@HoTenBo", hoTenBo);
                cmd.Parameters.AddWithValue("@NgaySinhBo", ngaySinhBo);
                cmd.Parameters.AddWithValue("@NgheNghiepBo", ngheNghiepBo);
                cmd.Parameters.AddWithValue("@SDTBo", sdtBo);

                cmd.Parameters.AddWithValue("@HoTenMe", hoTenMe);
                cmd.Parameters.AddWithValue("@NgaySinhMe", ngaySinhMe);
                cmd.Parameters.AddWithValue("@NgheNghiepMe", ngheNghiepMe);
                cmd.Parameters.AddWithValue("@SDTMe", sdtMe);

                cmd.Parameters.AddWithValue("@HoTenVo_Chong", hoTenVo);
                cmd.Parameters.AddWithValue("@SDTVo_Chong", sdtVo);
                cmd.Parameters.AddWithValue("@NgaySinhVo_Chong",
                    ngaySinhVo.HasValue
                    ? (object)ngaySinhVo.Value
                    : DBNull.Value);

                cmd.Parameters.AddWithValue("@NgheNghiepVo_Chong", ngheNghiepVo);

                cmd.Parameters.AddWithValue("@HoTenCon1", hoTenCon1);
                cmd.Parameters.AddWithValue("@NgaySinhCon1",
                    ngaySinhCon1.HasValue
                    ? (object)ngaySinhCon1.Value
                    : DBNull.Value);

                cmd.Parameters.AddWithValue("@HocVanCon1", hocVanCon1);

                cmd.Parameters.AddWithValue("@HoTenCon2", hoTenCon2);
                cmd.Parameters.AddWithValue("@NgaySinhCon2",
                    ngaySinhCon2.HasValue
                    ? (object)ngaySinhCon2.Value
                    : DBNull.Value);

                cmd.Parameters.AddWithValue("@HocVanCon2", hocVanCon2);

                cmd.Parameters.AddWithValue("@HoTenCon3", hoTenCon3);
                cmd.Parameters.AddWithValue("@NgaySinhCon3",
                    ngaySinhCon3.HasValue
                    ? (object)ngaySinhCon3.Value
                    : DBNull.Value);

                cmd.Parameters.AddWithValue("@HocVanCon3", hocVanCon3);

                cmd.Parameters.AddWithValue("@GhiChu", ghiChu);

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Sửa thông tin gia đình thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ConnectData.dongketnoi();
            }

            
            load_thongtingiadinh();
             Loadnhanvien();
             clearForm();
        }

        private void Loadnhanvien()
        {

            this.lku_nhanvien.Properties.DataSource = ConnectData.getdata("SELECT \r\n        nv.*,\r\n        CONCAT(nv.HoDem, ' ', nv.Ten) AS HoTen\r\n    FROM NHANVIEN nv\r\n    LEFT JOIN THONGTINGIADINH bh\r\n        ON nv.id = bh.NhanVien_id\r\n    WHERE nv.DELETEO_BY IS NULL\r\n        AND bh.NhanVien_id IS NULL\r\n    ORDER BY nv.id ASC");
            this.lku_nhanvien.Properties.DisplayMember = "HoTen";
            this.lku_nhanvien.Properties.ValueMember = "id";
        }
        private string VietHoaChuCaiDau(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = text.Trim().ToLower();

            string[] arr = text.Split(' ');

            for (int i = 0; i < arr.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(arr[i]))
                {
                    arr[i] =
                        char.ToUpper(arr[i][0]) +
                        arr[i].Substring(1);
                }
            }

            return string.Join(" ", arr);
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                // ===== Check nhân viên =====
                if (lku_nhanvien.EditValue == null)
                {
                    MessageBox.Show(
                        "Vui lòng chọn nhân viên");

                    lku_nhanvien.Focus();
                    return;
                }

                // ===== Check họ tên bố =====
                if (string.IsNullOrWhiteSpace(
                    txt_tenbo.Text))
                {
                    MessageBox.Show(
                        "Vui lòng nhập họ tên bố");

                    txt_tenbo.Focus();
                    return;
                }

                // ===== Check họ tên mẹ =====
                if (string.IsNullOrWhiteSpace(
                    txt_me.Text))
                {
                    MessageBox.Show(
                        "Vui lòng nhập họ tên mẹ");

                    txt_me.Focus();
                    return;
                }

                // ===== Check số điện thoại =====
                if (!string.IsNullOrWhiteSpace(
                    txt_sdtBo.Text)
                    && !txt_sdtBo.Text.All(char.IsDigit))
                {
                    MessageBox.Show(
                        "SĐT bố chỉ được nhập số");

                    txt_sdtBo.Focus();
                    return;
                }

                if (!string.IsNullOrWhiteSpace(
                    txt_sdtMe.Text)
                    && !txt_sdtMe.Text.All(char.IsDigit))
                {
                    MessageBox.Show(
                        "SĐT mẹ chỉ được nhập số");

                    txt_sdtMe.Focus();
                    return;
                }

                if (!string.IsNullOrWhiteSpace(
                    txt_sdtvo.Text)
                    && !txt_sdtvo.Text.All(char.IsDigit))
                {
                    MessageBox.Show(
                        "SĐT vợ/chồng chỉ được nhập số");

                    txt_sdtvo.Focus();
                    return;
                }

                // ===== Convert ngày sinh vợ/chồng =====
                DateTime ngaySinhVo = DateTime.MinValue;

                if (!string.IsNullOrWhiteSpace(
                    txt_ngaysinhvo.Text))
                {
                    bool checkNgayVo =
                        DateTime.TryParseExact(
                            txt_ngaysinhvo.Text.Trim(),
                            "dd/MM/yyyy",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out ngaySinhVo);

                    if (!checkNgayVo)
                    {
                        MessageBox.Show(
                            "Ngày sinh vợ/chồng không đúng định dạng dd/MM/yyyy");

                        txt_ngaysinhvo.Focus();
                        return;
                    }
                }

                // ===== Convert ngày sinh con 1 =====
                DateTime ngaySinhCon1 = DateTime.MinValue;

                if (!string.IsNullOrWhiteSpace(
                    date_con1.Text))
                {
                    bool checkCon1 =
                        DateTime.TryParseExact(
                            date_con1.Text.Trim(),
                            "dd/MM/yyyy",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out ngaySinhCon1);

                    if (!checkCon1)
                    {
                        MessageBox.Show(
                            "Ngày sinh con 1 không đúng định dạng dd/MM/yyyy");

                        date_con1.Focus();
                        return;
                    }
                }

                // ===== Convert ngày sinh con 2 =====
                DateTime ngaySinhCon2 = DateTime.MinValue;

                if (!string.IsNullOrWhiteSpace(
                    date_con2.Text))
                {
                    bool checkCon2 =
                        DateTime.TryParseExact(
                            date_con2.Text.Trim(),
                            "dd/MM/yyyy",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out ngaySinhCon2);

                    if (!checkCon2)
                    {
                        MessageBox.Show(
                            "Ngày sinh con 2 không đúng định dạng dd/MM/yyyy");

                        date_con2.Focus();
                        return;
                    }
                }

                // ===== Convert ngày sinh con 3 =====
                DateTime ngaySinhCon3 = DateTime.MinValue;

                if (!string.IsNullOrWhiteSpace(
                    date_con3.Text))
                {
                    bool checkCon3 =
                        DateTime.TryParseExact(
                            date_con3.Text.Trim(),
                            "dd/MM/yyyy",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out ngaySinhCon3);

                    if (!checkCon3)
                    {
                        MessageBox.Show(
                            "Ngày sinh con 3 không đúng định dạng dd/MM/yyyy");

                        date_con3.Focus();
                        return;
                    }
                }

                ConnectData.taoketnoi();

                int nhanVienID =
                    Convert.ToInt32(
                        lku_nhanvien.EditValue);

                // ===== Check đã tồn tại =====
                string sqlCheck =
                    @"SELECT COUNT(*)
              FROM THONGTINGIADINH
              WHERE NhanVien_id=@NhanVien_id";

                MySqlCommand cmdCheck =
                    new MySqlCommand(
                        sqlCheck,
                        ConnectData.conn);

                cmdCheck.Parameters.AddWithValue(
                    "@NhanVien_id",
                    nhanVienID);

                int check =
                    Convert.ToInt32(
                        cmdCheck.ExecuteScalar());

                if (check > 0)
                {
                    MessageBox.Show(
                        "Nhân viên này đã có thông tin gia đình");

                    return;
                }

                // ===== SQL INSERT =====
                string sql = @"
INSERT INTO THONGTINGIADINH
(
    NhanVien_id,

    HoTenBo,
    NgaySinhBo,
    NgheNghiepBo,
    SDTBo,

    HoTenMe,
    NgaySinhMe,
    NgheNghiepMe,
    SDTMe,

    HoTenVo_Chong,
    SDTVo_Chong,
    NgaySinhVo_Chong,
    NgheNghiepVo_Chong,

    HoTenCon1,
    NgaySinhCon1,
    HocVanCon1,

    HoTenCon2,
    NgaySinhCon2,
    HocVanCon2,

    HoTenCon3,
    NgaySinhCon3,
    HocVanCon3,

    GhiChu
)
VALUES
(
    @NhanVien_id,

    @HoTenBo,
    @NgaySinhBo,
    @NgheNghiepBo,
    @SDTBo,

    @HoTenMe,
    @NgaySinhMe,
    @NgheNghiepMe,
    @SDTMe,

    @HoTenVo_Chong,
    @SDTVo_Chong,
    @NgaySinhVo_Chong,
    @NgheNghiepVo_Chong,

    @HoTenCon1,
    @NgaySinhCon1,
    @HocVanCon1,

    @HoTenCon2,
    @NgaySinhCon2,
    @HocVanCon2,

    @HoTenCon3,
    @NgaySinhCon3,
    @HocVanCon3,

    @GhiChu
)";

                MySqlCommand cmd =
                    new MySqlCommand(
                        sql,
                        ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@NhanVien_id",
                    nhanVienID);

                // ===== Bố =====
                cmd.Parameters.AddWithValue(
                    "@HoTenBo",
                    VietHoaChuCaiDau(
                        txt_tenbo.Text));

                cmd.Parameters.AddWithValue(
                    "@NgaySinhBo",
                    date_NgaySinhBo.Value.Date);

                cmd.Parameters.AddWithValue(
                    "@NgheNghiepBo",
                    txt_nghenghiepbo.Text.Trim());

                cmd.Parameters.AddWithValue(
                    "@SDTBo",
                    txt_sdtBo.Text.Trim());

                // ===== Mẹ =====
                cmd.Parameters.AddWithValue(
                    "@HoTenMe",
                    VietHoaChuCaiDau(
                        txt_me.Text));

                cmd.Parameters.AddWithValue(
                    "@NgaySinhMe",
                    date_me.Value.Date);

                cmd.Parameters.AddWithValue(
                    "@NgheNghiepMe",
                    txt_NgheNghiepMe.Text.Trim());

                cmd.Parameters.AddWithValue(
                    "@SDTMe",
                    txt_sdtMe.Text.Trim());

                // ===== Vợ / Chồng =====
                cmd.Parameters.AddWithValue(
                    "@HoTenVo_Chong",
                    VietHoaChuCaiDau(
                        txt_voChong.Text));

                cmd.Parameters.AddWithValue(
                    "@SDTVo_Chong",
                    txt_sdtvo.Text.Trim());

                if (string.IsNullOrWhiteSpace(
                    txt_ngaysinhvo.Text))
                {
                    cmd.Parameters.AddWithValue(
                        "@NgaySinhVo_Chong",
                        DBNull.Value);
                }
                else
                {
                    cmd.Parameters.AddWithValue(
                        "@NgaySinhVo_Chong",
                        ngaySinhVo.ToString("yyyy-MM-dd"));
                }

                cmd.Parameters.AddWithValue(
                    "@NgheNghiepVo_Chong",
                    txt_ngheNghiepVo.Text.Trim());

                // ===== Con 1 =====
                cmd.Parameters.AddWithValue(
                    "@HoTenCon1",
                    VietHoaChuCaiDau(
                        txt_tencon1.Text));

                if (string.IsNullOrWhiteSpace(
                    date_con1.Text))
                {
                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon1",
                        DBNull.Value);
                }
                else
                {
                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon1",
                        ngaySinhCon1.ToString("yyyy-MM-dd"));
                }

                cmd.Parameters.AddWithValue(
                    "@HocVanCon1",
                    cb_hocvan1.Text.Trim());

                // ===== Con 2 =====
                cmd.Parameters.AddWithValue(
                    "@HoTenCon2",
                    VietHoaChuCaiDau(
                        txt_tencon2.Text));

                if (string.IsNullOrWhiteSpace(
                    date_con2.Text))
                {
                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon2",
                        DBNull.Value);
                }
                else
                {
                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon2",
                        ngaySinhCon2.ToString("yyyy-MM-dd"));
                }

                cmd.Parameters.AddWithValue(
                    "@HocVanCon2",
                    cb_hocvan2.Text.Trim());

                // ===== Con 3 =====
                cmd.Parameters.AddWithValue(
                    "@HoTenCon3",
                    VietHoaChuCaiDau(
                        txt_tencon3.Text));

                if (string.IsNullOrWhiteSpace(
                    date_con3.Text))
                {
                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon3",
                        DBNull.Value);
                }
                else
                {
                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon3",
                        ngaySinhCon3.ToString("yyyy-MM-dd"));
                }

                cmd.Parameters.AddWithValue(
                    "@HocVanCon3",
                    cb_hocvan3.Text.Trim());

                // ===== Ghi chú =====
                cmd.Parameters.AddWithValue(
                    "@GhiChu",
                    txt_ghichu.Text.Trim());

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Thêm thông tin gia đình thành công",
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
            load_thongtingiadinh();
            Loadnhanvien();
            clearForm();

        }
        void clearForm()
        {
            lku_nhanvien.EditValue = null;
            txt_tenbo.Text = "";
            date_NgaySinhBo.Value = DateTime.Now;
            txt_nghenghiepbo.Text = "";
            txt_sdtBo.Text = "";
            txt_me.Text = "";
            date_me.Value = DateTime.Now;
            txt_NgheNghiepMe.Text = "";
            txt_sdtMe.Text = "";
            txt_voChong.Text = "";
            txt_sdtvo.Text = "";
            txt_ngaysinhvo.Text = "";
            txt_ngheNghiepVo.Text = "";
            txt_tencon1.Text = "";
            date_con1.Text = "";
            cb_hocvan1.SelectedIndex = -1;
            txt_tencon2.Text = "";
            date_con2.Text = "";
            cb_hocvan2.SelectedIndex = -1;
            txt_tencon3.Text = "";
            date_con3.Text = "";
            cb_hocvan3.SelectedIndex = -1;
            txt_ghichu.Text = "";
        }
        private void load_thongtingiadinh()
        {
            gc_giadinh.DataSource = ConnectData.getdata("SELECT \r\n    gd.Id,\r\n    nv.MaNV,\r\n    CONCAT(nv.HoDem, ' ', nv.Ten) AS HoTen,\r\n    nv.NgaySinh,\r\n    nv.SDT,\r\n    nv.Email,\r\n\r\n    gd.HoTenBo,\r\n    gd.NgaySinhBo,\r\n    gd.NgheNghiepBo,\r\n    gd.SDTBo,\r\n\r\n    gd.HoTenMe,\r\n    gd.NgaySinhMe,\r\n    gd.NgheNghiepMe,\r\n    gd.SDTMe,\r\n\r\n    gd.HoTenVo_Chong,\r\n    gd.SDTVo_Chong,\r\n    gd.NgaySinhVo_Chong,\r\n    gd.NgheNghiepVo_Chong,\r\n\r\n    gd.HoTenCon1,\r\n    gd.NgaySinhCon1,\r\n    gd.HocVanCon1,\r\n\r\n    gd.HoTenCon2,\r\n    gd.NgaySinhCon2,\r\n    gd.HocVanCon2,\r\n\r\n    gd.HoTenCon3,\r\n    gd.NgaySinhCon3,\r\n    gd.HocVanCon3,\r\n\r\n    gd.GhiChu\r\n\r\nFROM THONGTINGIADINH gd\r\n\r\nINNER JOIN NHANVIEN nv\r\n    ON nv.id = gd.NhanVien_id\r\n\r\nWHERE nv.DELETEO_BY IS NULL\r\n\r\nORDER BY gd.Id DESC");

        }

        private void txt_sdtBo_KeyPress(object sender, KeyPressEventArgs e)
        {
            // ===== Chỉ cho nhập số =====
            if (!char.IsControl(e.KeyChar)
                && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txt_sdtMe_KeyPress(object sender, KeyPressEventArgs e)
        {
            // ===== Chỉ cho nhập số =====
            if (!char.IsControl(e.KeyChar)
                && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txt_sdtvo_KeyPress(object sender, KeyPressEventArgs e)
        {
            // ===== Chỉ cho nhập số =====
            if (!char.IsControl(e.KeyChar)
                && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void btn_lammoi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            load_thongtingiadinh();
            Loadnhanvien();
            clearForm();
        }

        private void btn_xuat_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                SaveFileDialog save =
                    new SaveFileDialog();

                save.Filter =
                    "Excel Workbook|*.xlsx";

                save.Title =
                    "Xuất Excel";

                save.FileName =
                    "DanhSachThongTinGiaDinh.xlsx";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                Microsoft.Office.Interop.Excel.Application app =
                    new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook wb =
                    app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet ws =
                    (Microsoft.Office.Interop.Excel.Worksheet)
                    wb.ActiveSheet;

                ws.Name = "THONGTINGIADINH";

                int cot =
                    gv_giadinh.Columns.Count;

                int dong =
                    gv_giadinh.RowCount;

                // ===== TIÊU ĐỀ =====
                Microsoft.Office.Interop.Excel.Range title =
                    ws.Range["A1", "AC1"];

                title.Merge();

                title.Value =
                    "DANH SÁCH THÔNG TIN GIA ĐÌNH NHÂN VIÊN";

                title.Font.Bold = true;

                title.Font.Size = 18;

                title.HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // ===== HEADER =====
                int colExcel = 1;

                for (int i = 0; i < cot; i++)
                {
                    string tenCot =
                        gv_giadinh.Columns[i].FieldName;

                    // ===== Bỏ cột ID =====
                    if (tenCot == "Id")
                        continue;

                    ws.Cells[3, colExcel] =
                        gv_giadinh.Columns[i].Caption;

                    ws.Cells[3, colExcel].Font.Bold =
                        true;

                    ws.Cells[3, colExcel].Borders.LineStyle =
                        1;

                    ws.Cells[3, colExcel].Interior.Color =
                        System.Drawing.Color.LightGray;

                    colExcel++;
                }

                // ===== DỮ LIỆU =====
                for (int i = 0; i < dong; i++)
                {
                    colExcel = 1;

                    for (int j = 0; j < cot; j++)
                    {
                        string tenCot =
                            gv_giadinh.Columns[j].FieldName;

                        if (tenCot == "Id")
                            continue;

                        object value =
                            gv_giadinh.GetRowCellValue(
                                i,
                                gv_giadinh.Columns[j]);

                        // ===== Format ngày =====
                        if (value != null &&
                           (
                            tenCot == "NgaySinh"
                            || tenCot == "NgaySinhBo"
                            || tenCot == "NgaySinhMe"
                            || tenCot == "NgaySinhVo_Chong"
                            || tenCot == "NgaySinhCon1"
                            || tenCot == "NgaySinhCon2"
                            || tenCot == "NgaySinhCon3"
                           ))
                        {
                            DateTime ngay;

                            if (DateTime.TryParse(
                                value.ToString(),
                                out ngay))
                            {
                                ws.Cells[i + 4, colExcel] =
                                    ngay.ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                ws.Cells[i + 4, colExcel] =
                                    value.ToString();
                            }
                        }
                        else
                        {
                            ws.Cells[i + 4, colExcel] =
                                value?.ToString();
                        }

                        // ===== Border =====
                        ws.Cells[i + 4, colExcel]
                            .Borders.LineStyle = 1;

                        colExcel++;
                    }
                }

                // ===== Căn giữa =====
                ws.Range["A3", ws.Cells[dong + 3, colExcel - 1]]
                    .HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // ===== AutoFit =====
                ws.Columns.AutoFit();

                // ===== Lưu =====
                wb.SaveAs(save.FileName);

                wb.Close();

                app.Quit();

                MessageBox.Show(
                    "Xuất Excel thành công",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_import_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                OpenFileDialog open =
                    new OpenFileDialog();

                open.Filter =
                    "Excel File|*.xlsx;*.xls";

                if (open.ShowDialog() != DialogResult.OK)
                    return;

                Excel.Application app =
                    new Excel.Application();

                Excel.Workbook wb =
                    app.Workbooks.Open(open.FileName);

                Excel.Worksheet ws =
                    (Excel.Worksheet)wb.Sheets[1];

                Excel.Range range =
                    ws.UsedRange;

                int rowCount =
                    range.Rows.Count;

                ConnectData.taoketnoi();

                // ===== Bắt đầu từ dòng 4 =====
                for (int i = 4; i <= rowCount; i++)
                {
                    // ===== Mã nhân viên =====
                    string maNV =
                        Convert.ToString(
                            (range.Cells[i, 2] as Excel.Range).Text);

                    if (string.IsNullOrWhiteSpace(maNV))
                        continue;

                    // ===== Tìm nhân viên =====
                    string sqlNhanVien =
                        @"SELECT id
                  FROM NHANVIEN
                  WHERE MaNV=@MaNV
                  AND DELETEO_BY IS NULL";

                    MySqlCommand cmdNhanVien =
                        new MySqlCommand(
                            sqlNhanVien,
                            ConnectData.conn);

                    cmdNhanVien.Parameters.AddWithValue(
                        "@MaNV",
                        maNV);

                    object result =
                        cmdNhanVien.ExecuteScalar();

                    if (result == null)
                    {
                        MessageBox.Show(
                            "Không tìm thấy nhân viên mã: " + maNV);

                        continue;
                    }

                    int nhanVienID =
                        Convert.ToInt32(result);

                    // ===== Check tồn tại =====
                    string sqlCheck =
                        @"SELECT COUNT(*)
                  FROM THONGTINGIADINH
                  WHERE NhanVien_id=@NhanVien_id";

                    MySqlCommand cmdCheck =
                        new MySqlCommand(
                            sqlCheck,
                            ConnectData.conn);

                    cmdCheck.Parameters.AddWithValue(
                        "@NhanVien_id",
                        nhanVienID);

                    int check =
                        Convert.ToInt32(
                            cmdCheck.ExecuteScalar());

                    if (check > 0)
                    {
                        MessageBox.Show(
                            "Nhân viên mã "
                            + maNV
                            + " đã có thông tin gia đình");

                        continue;
                    }

                    // ===== Hàm lấy text =====
                    string GetText(int col)
                    {
                        return Convert.ToString(
                            (range.Cells[i, col] as Excel.Range).Text).Trim();
                    }

                    // ===== Hàm convert ngày =====
                    object ConvertDate(int col)
                    {
                        try
                        {
                            object value =
                                (range.Cells[i, col] as Excel.Range).Value2;

                            if (value == null)
                                return DBNull.Value;

                            double oaDate;

                            if (double.TryParse(
                                value.ToString(),
                                out oaDate))
                            {
                                return DateTime
                                    .FromOADate(oaDate)
                                    .ToString("yyyy-MM-dd");
                            }

                            DateTime ngay;

                            bool checkNgay =
                                DateTime.TryParseExact(
                                    value.ToString(),
                                    "dd/MM/yyyy",
                                    CultureInfo.InvariantCulture,
                                    DateTimeStyles.None,
                                    out ngay);

                            if (checkNgay)
                            {
                                return ngay.ToString("yyyy-MM-dd");
                            }

                            return DBNull.Value;
                        }
                        catch
                        {
                            return DBNull.Value;
                        }
                    }

                    // ===== INSERT =====
                    string sql = @"
                        INSERT INTO THONGTINGIADINH
                        (
                            NhanVien_id,

                            HoTenBo,
                            NgaySinhBo,
                            NgheNghiepBo,
                            SDTBo,

                            HoTenMe,
                            NgaySinhMe,
                            NgheNghiepMe,
                            SDTMe,

                            HoTenVo_Chong,
                            NgaySinhVo_Chong,
                            SDTVo_Chong,
                            NgheNghiepVo_Chong,

                            HoTenCon1,
                            NgaySinhCon1,
                            HocVanCon1,

                            HoTenCon2,
                            NgaySinhCon2,
                            HocVanCon2,

                            HoTenCon3,
                            NgaySinhCon3,
                            HocVanCon3,

                            GhiChu
                        )
                        VALUES
                        (
                            @NhanVien_id,

                            @HoTenBo,
                            @NgaySinhBo,
                            @NgheNghiepBo,
                            @SDTBo,

                            @HoTenMe,
                            @NgaySinhMe,
                            @NgheNghiepMe,
                            @SDTMe,

                            @HoTenVo_Chong,
                            @NgaySinhVo_Chong,
                            @SDTVo_Chong,
                            @NgheNghiepVo_Chong,

                            @HoTenCon1,
                            @NgaySinhCon1,
                            @HocVanCon1,

                            @HoTenCon2,
                            @NgaySinhCon2,
                            @HocVanCon2,

                            @HoTenCon3,
                            @NgaySinhCon3,
                            @HocVanCon3,

                            @GhiChu
                        )";

                    MySqlCommand cmd =
                        new MySqlCommand(
                            sql,
                            ConnectData.conn);

                    cmd.Parameters.AddWithValue(
                        "@NhanVien_id",
                        nhanVienID);

                    // ===== Bố =====
                    cmd.Parameters.AddWithValue(
                        "@HoTenBo",
                        GetText(7));

                    cmd.Parameters.AddWithValue(
                        "@NgaySinhBo",
                        ConvertDate(8));

                    cmd.Parameters.AddWithValue(
                        "@NgheNghiepBo",
                        GetText(9));

                    cmd.Parameters.AddWithValue(
                        "@SDTBo",
                        GetText(10));

                    // ===== Mẹ =====
                    cmd.Parameters.AddWithValue(
                        "@HoTenMe",
                        GetText(11));

                    cmd.Parameters.AddWithValue(
                        "@NgaySinhMe",
                        ConvertDate(12));

                    cmd.Parameters.AddWithValue(
                        "@NgheNghiepMe",
                        GetText(13));

                    cmd.Parameters.AddWithValue(
                        "@SDTMe",
                        GetText(14));

                    // ===== Vợ chồng =====
                    cmd.Parameters.AddWithValue(
                        "@HoTenVo_Chong",
                        GetText(15));

                    cmd.Parameters.AddWithValue(
                        "@NgaySinhVo_Chong",
                        ConvertDate(16));

                    cmd.Parameters.AddWithValue(
                        "@SDTVo_Chong",
                        GetText(17));

                    cmd.Parameters.AddWithValue(
                        "@NgheNghiepVo_Chong",
                        GetText(18));

                    // ===== Con 1 =====
                    cmd.Parameters.AddWithValue(
                        "@HoTenCon1",
                        GetText(19));

                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon1",
                        ConvertDate(20));

                    cmd.Parameters.AddWithValue(
                        "@HocVanCon1",
                        GetText(21));

                    // ===== Con 2 =====
                    cmd.Parameters.AddWithValue(
                        "@HoTenCon2",
                        GetText(22));

                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon2",
                        ConvertDate(23));

                    cmd.Parameters.AddWithValue(
                        "@HocVanCon2",
                        GetText(24));

                    // ===== Con 3 =====
                    cmd.Parameters.AddWithValue(
                        "@HoTenCon3",
                        GetText(25));

                    cmd.Parameters.AddWithValue(
                        "@NgaySinhCon3",
                        ConvertDate(26));

                    cmd.Parameters.AddWithValue(
                        "@HocVanCon3",
                        GetText(27));

                    // ===== Ghi chú =====
                    cmd.Parameters.AddWithValue(
                        "@GhiChu",
                        GetText(28));

                    cmd.ExecuteNonQuery();
                }

                wb.Close(false);
                app.Quit();

                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);

                MessageBox.Show(
                    "Import Excel thành công",
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

            load_thongtingiadinh();
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