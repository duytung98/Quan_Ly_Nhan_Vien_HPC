using ClosedXML.Excel;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class F_HopDong1 : DevExpress.XtraEditors.XtraForm
    {
        public F_HopDong1()
        {
            InitializeComponent();
        }

        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                // ===== KIỂM TRA TRỐNG =====
                if (lku_nhanvien.EditValue == null)
                {
                    MessageBox.Show(
                        "Vui lòng chọn nhân viên");

                    lku_nhanvien.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(
                    cb_loaihopdong.Text))
                {
                    MessageBox.Show(
                        "Vui lòng chọn loại hợp đồng");

                    cb_loaihopdong.Focus();
                    return;
                }

                if (lbl_phongban.EditValue == null)
                {
                    MessageBox.Show(
                        "Vui lòng chọn phòng ban");

                    lbl_phongban.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(
                    txt_chucvu.Text))
                {
                    MessageBox.Show(
                        "Vui lòng nhập chức vụ");

                    txt_chucvu.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(
                    txt_luongcb.Text))
                {
                    MessageBox.Show(
                        "Vui lòng nhập lương cơ bản");

                    txt_luongcb.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(
                    txt_hesoluong.Text))
                {
                    MessageBox.Show(
                        "Vui lòng nhập hệ số lương");

                    txt_hesoluong.Focus();
                    return;
                }

                ConnectData.taoketnoi();

                int nhanVienID =
                    Convert.ToInt32(
                        lku_nhanvien.EditValue);

                // ===== CHECK MỖI NHÂN VIÊN 1 HỢP ĐỒNG =====
                string sqlCheck =
                    "SELECT COUNT(*) FROM HOPDONG WHERE NhanVien_id=@NhanVien_id";

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
                        "Nhân viên này đã có hợp đồng");

                    return;
                }

                // ===== TẠO SỐ HỢP ĐỒNG =====
                int stt = 1;

                string nam =
                    DateTime.Now.Year.ToString();

                string sqlSoHD =
                    @"SELECT COUNT(*) + 1
          FROM HOPDONG
          WHERE YEAR(NgayBD)=YEAR(NOW())";

                MySqlCommand cmdSoHD =
                    new MySqlCommand(
                        sqlSoHD,
                        ConnectData.conn);

                stt =
                    Convert.ToInt32(
                        cmdSoHD.ExecuteScalar());

                string soHopDong =
                    stt.ToString("0000")
                    + "/"
                    + nam
                    + "/HĐLĐ";


                // ===== Convert ngày sinh vợ/chồng =====
                DateTime ngayKT = DateTime.MinValue;

                if (!string.IsNullOrWhiteSpace(
                    txt_ngayKT.Text))
                {
                    bool checkNgayKT =
                        DateTime.TryParseExact(
                            txt_ngayKT.Text.Trim(),
                            "dd/MM/yyyy",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out ngayKT);        
                    if (!checkNgayKT)
                    {
                        MessageBox.Show(
                            "Ngày kết thúc không đúng định dạng dd/MM/yyyy");

                        txt_ngayKT.Focus();
                        return;
                    }
                }
                

                // ===== INSERT =====
                string sql = @"
    INSERT INTO HOPDONG
    (
        NhanVien_id,
        So_HopDong,
        NgayBD,
        NgayKT,
        LoaiHopDong,
        PhongBan_ID,
        ChucVu,
        LuongCB,
        HeSoLuong,
        GhiChu
    )
    VALUES
    (
        @NhanVien_id,
        @So_HopDong,
        @NgayBD,
        @NgayKT,
        @LoaiHopDong,
        @PhongBan_ID,
        @ChucVu,
        @LuongCB,
        @HeSoLuong,
        @GhiChu
    )";

                MySqlCommand cmd =
                    new MySqlCommand(
                        sql,
                        ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@NhanVien_id",
                    nhanVienID);

                cmd.Parameters.AddWithValue(
                    "@So_HopDong",
                    soHopDong);

                cmd.Parameters.AddWithValue(
                    "@NgayBD",
                    date_NgayBD.Value.Date);

                cmd.Parameters.AddWithValue(
                    "@NgayKT",
                    ngayKT);

                cmd.Parameters.AddWithValue(
                    "@LoaiHopDong",
                    cb_loaihopdong.Text);

                cmd.Parameters.AddWithValue(
                    "@PhongBan_ID",
                    lbl_phongban.EditValue);

                cmd.Parameters.AddWithValue(
                    "@ChucVu",
                    txt_chucvu.Text.Trim());

                cmd.Parameters.AddWithValue(
                    "@LuongCB",
                    Convert.ToInt32(
                        txt_luongcb.Text
                        .Replace(".", "")
                        .Replace(",", "")));

                cmd.Parameters.AddWithValue(
                    "@HeSoLuong",
                    Convert.ToSingle(
                        txt_hesoluong.Text
                        .Replace(",", ".")));

                cmd.Parameters.AddWithValue(
                    "@GhiChu",
                    txt_ghichu.Text.Trim());

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Thêm hợp đồng thành công");

                LogSystem.WriteLog(
                    "Hợp đồng",
                    "THÊM",
                    "Thêm hợp đồng cho nhân viên ID: "
                    + nhanVienID
                    + " - "
                    + HoTen);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ConnectData.dongketnoi();
            }

            load_hopdong();

            clearForm();

            Loadnhanvien();
        }
        DangNhap Login = (DangNhap)Application.OpenForms["DangNhap"];
        private void btn_lammoi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            load_hopdong();
            clearForm();
            Loadnhanvien();
        }
        void load_lkuPhongban()
        {
            this.lbl_phgban1.DataSource = ConnectData.getdata("SELECT * FROM PHONGBAN");
            this.lbl_phgban1.DisplayMember = "TenPhongBan";
            this.lbl_phgban1.ValueMember = "id";

        }

        private void F_HopDong1_Load(object sender, EventArgs e)
        {
            load_lkuPhongban();
            lku_nhanvien.Focus();
            cb_loaihopdong.Text = "Hợp đồng dài hạn";
            Loadnhanvien();
            load_hopdong();
            Load_lbl_phongban();
            gv_hopdong.OptionsBehavior.Editable = true;
            gv_hopdong.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None;
            GridView gridView = gv_hopdong;
            // Định dạng căn giữa chữ trong Header
            gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            // Định dạng căn giữa chữ trong Row
            gridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            btn_sua.Click += Btn_sua_Click;
            btn_xoa.Click += Btn_xoa_Click;
            lku_nhanvien.KeyDown += lku_nhanvien_KeyDown;
            txt_luongcb.KeyDown += txt_luongcb_KeyDown;
            txt_ghichu.KeyDown += txt_ghichu_KeyDown;
            txt_hesoluong.KeyDown += txt_hesoluong_KeyDown;
            cb_loaihopdong.KeyDown += cb_loaihopdong_KeyDown;
            date_NgayBD.KeyDown += date_NgayBD_KeyDown;
            txt_chucvu.KeyDown += txt_chucvu_KeyDown;
            lbl_phongban.KeyDown += lbl_phongban_KeyDown;
            if (gv_hopdong.Columns["STT"] != null)
                gv_hopdong.Columns["STT"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_hopdong.Columns["MaNV"] != null)
                gv_hopdong.Columns["MaNV"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_hopdong.Columns["HoTen"] != null)
                gv_hopdong.Columns["HoTen"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_hopdong.Columns["btn_xoa"] != null)
                gv_hopdong.Columns["btn_xoa"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;

            if (gv_hopdong.Columns["btn_sua"] != null)
                gv_hopdong.Columns["btn_sua"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;



            // ===== Tự động độ rộng =====
            gv_hopdong.OptionsView.ColumnAutoWidth = false;
            gv_hopdong.BestFitColumns();

            // ===== Thanh cuộn =====
            gv_hopdong.HorzScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;

            gv_hopdong.VertScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            gv_hopdong.Columns["HoTen"].Width = 200;

            txt_ngayKT.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTime;
            txt_ngayKT.Properties.Mask.EditMask = "dd/MM/yyyy";
            txt_ngayKT.Properties.Mask.UseMaskAsDisplayFormat = true;
        }

        private void lbl_phongban_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_luongcb.Focus();
            }
        }

        private void txt_chucvu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                lbl_phongban.Focus();
            }
        }

        private void txt_hesoluong_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_ghichu.Focus();
            }
        }

        private void txt_ghichu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_them.PerformClick();
            }
        }

        private void txt_luongcb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_hesoluong.Focus();
            }
        }

        private void cb_loaihopdong_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_chucvu.Focus();
            }
        }

        private void date_NgayBD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                cb_loaihopdong.Focus();
            }
        }

        private void lku_nhanvien_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                date_NgayBD.Focus();
            }
        }

        private void Btn_xoa_Click(object sender, EventArgs e)
        {
            if (gv_hopdong.FocusedRowHandle < 0)
            {
                MessageBox.Show("Chọn hợp đồng cần xóa");
                return;
            }

            string hoTen =
                gv_hopdong.GetFocusedRowCellValue("HoTen")?.ToString();

            string soHopDong =
                gv_hopdong.GetFocusedRowCellValue("So_HopDong")?.ToString();

            DialogResult rs = MessageBox.Show(
                "Bạn có chắc chắn muốn xóa hợp đồng:\n"
                + soHopDong
                + "\n\nCủa nhân viên:\n"
                + hoTen
                + " ?",
                "Thông báo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (rs == DialogResult.No)
                return;

            try
            {
                ConnectData.taoketnoi();

                int id =
                    Convert.ToInt32(
                        gv_hopdong.GetFocusedRowCellValue("Id"));

                string sql =
                    "DELETE FROM HOPDONG WHERE Id=@Id";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@Id",
                    id);

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Xóa hợp đồng thành công của:\n"
                    + hoTen,
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                LogSystem.WriteLog("Nhân viên", "XÓA", "Xóa hợp đồng: " + soHopDong + " của nhân viên ID: " + MaNV + " - " + HoTen + " - " + Login.txt_taikhoan.Text);


            }
            catch (MySqlException ex)
            {
                if (ex.Number == 1451)
                {
                    MessageBox.Show(
                        "Không thể xóa vì dữ liệu đang được sử dụng",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
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
            load_hopdong();
            Loadnhanvien();
            clearForm();
        }

        private void Btn_sua_Click(object sender, EventArgs e)
        {
            try
            {
                if (gv_hopdong.FocusedRowHandle < 0)
                    return;

                DialogResult rs =
                    MessageBox.Show(
                        "Bạn có chắc chắn muốn sửa hợp đồng này không?",
                        "Thông báo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                if (rs != DialogResult.Yes)
                    return;

                int id =
                    Convert.ToInt32(
                        gv_hopdong.GetFocusedRowCellValue("Id"));

                // ===== LẤY DỮ LIỆU TỪ GRIDVIEW =====
                string maNV =
                    Convert.ToString(
                        gv_hopdong.GetFocusedRowCellValue("MaNV"));

                string hoTen =
                    Convert.ToString(
                        gv_hopdong.GetFocusedRowCellValue("HoTen"));

                string loaiHopDong =
                    Convert.ToString(
                        gv_hopdong.GetFocusedRowCellValue("LoaiHopDong"));

                string chucVu =
                    Convert.ToString(
                        gv_hopdong.GetFocusedRowCellValue("ChucVu"));

                int phongBanID =
                    Convert.ToInt32(
                        gv_hopdong.GetFocusedRowCellValue("PhongBan_ID"));

                DateTime ngayBD =
                    Convert.ToDateTime(
                        gv_hopdong.GetFocusedRowCellValue("NgayBD"));

                // ===== NGÀY KẾT THÚC =====
                object ngayKT = DBNull.Value;

                object valueNgayKT =
                    gv_hopdong.GetFocusedRowCellValue("NgayKT");

                if (valueNgayKT != null &&
                    valueNgayKT != DBNull.Value)
                {
                    ngayKT =
                        Convert.ToDateTime(valueNgayKT);
                }

                int luongCB =
                    Convert.ToInt32(
                        gv_hopdong.GetFocusedRowCellValue("LuongCB"));

                float heSoLuong =
                    Convert.ToSingle(
                        gv_hopdong.GetFocusedRowCellValue("HeSoLuong"));

                string ghiChu =
                    Convert.ToString(
                        gv_hopdong.GetFocusedRowCellValue("GhiChu"));

                ConnectData.taoketnoi();

                string sql = @"
    UPDATE HOPDONG
    SET
        NgayBD=@NgayBD,
        NgayKT=@NgayKT,
        LoaiHopDong=@LoaiHopDong,
        PhongBan_ID=@PhongBan_ID,
        ChucVu=@ChucVu,
        LuongCB=@LuongCB,
        HeSoLuong=@HeSoLuong,
        GhiChu=@GhiChu
    WHERE Id=@Id";

                MySqlCommand cmd =
                    new MySqlCommand(
                        sql,
                        ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@NgayBD",
                    ngayBD);

                cmd.Parameters.AddWithValue(
                    "@NgayKT",
                    ngayKT);

                cmd.Parameters.AddWithValue(
                    "@LoaiHopDong",
                    loaiHopDong);

                cmd.Parameters.AddWithValue(
                    "@PhongBan_ID",
                    phongBanID);

                cmd.Parameters.AddWithValue(
                    "@ChucVu",
                    chucVu);

                cmd.Parameters.AddWithValue(
                    "@LuongCB",
                    luongCB);

                cmd.Parameters.AddWithValue(
                    "@HeSoLuong",
                    heSoLuong);

                cmd.Parameters.AddWithValue(
                    "@GhiChu",
                    ghiChu);

                cmd.Parameters.AddWithValue(
                    "@Id",
                    id);

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Sửa hợp đồng thành công");

                LogSystem.WriteLog(
                    "Nhân viên",
                    "SỬA",
                    "Sửa hợp đồng cho nhân viên: "
                    + maNV
                    + " - "
                    + hoTen
                    + " - "
                    + Login.txt_taikhoan.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ConnectData.dongketnoi();
            }

            load_hopdong();
        }
        
        void Loadnhanvien()
        {
            this.lku_nhanvien.Properties.DataSource = ConnectData.getdata("SELECT \r\n        nv.*,\r\n        CONCAT(nv.HoDem, ' ', nv.Ten) AS HoTen\r\n    FROM NHANVIEN nv\r\n    LEFT JOIN HOPDONG bh\r\n        ON nv.id = bh.NhanVien_id\r\n    WHERE nv.DELETEO_BY IS NULL\r\n        AND bh.NhanVien_id IS NULL\r\n    ORDER BY nv.id ASC");
            this.lku_nhanvien.Properties.DisplayMember = "HoTen";
            this.lku_nhanvien.Properties.ValueMember = "id";

        }
        void Load_lbl_phongban()
        {
            this.lbl_phongban.Properties.DataSource = ConnectData.getdata("SELECT * FROM PHONGBAN");
            this.lbl_phongban.Properties.DisplayMember = "TenPhongBan";
            this.lbl_phongban.Properties.ValueMember = "id";
        }

        private void load_hopdong()
        {
            gc_hopdong.DataSource = ConnectData.getdata("SELECT \r\n    hd.Id,\r\n    nv.MaNV,\r\n    CONCAT(nv.HoDem, ' ', nv.Ten) AS HoTen,\r\n    NgaySinh, SDT, Email,\r\n    hd.So_HopDong,\r\n    hd.NgayBD,hd.NgayKT,hd.PhongBan_ID,hd.ChucVu,\r\n    hd.LoaiHopDong,\r\n    hd.LuongCB,\r\n    hd.HeSoLuong,\r\n    ROUND(LuongCB * HeSoLuong,0) AS LuongThucNhan,\r\n    hd.GhiChu\r\nFROM HOPDONG hd\r\nINNER JOIN NHANVIEN nv\r\n    ON hd.NhanVien_id = nv.id\r\nWHERE nv.DELETEO_BY IS NULL\r\nORDER BY hd.Id desc;");

        }
        void clearForm()
        {
            lku_nhanvien.EditValue = null;
            cb_loaihopdong.Text = "Hợp đồng dài hạn";
            txt_luongcb.Text = "";
            txt_hesoluong.Text = "";
            txt_ghichu.Text = "";
            lbl_phongban.EditValue = null;
            txt_chucvu.Text ="";
            date_NgayBD.Value = DateTime.Now;
        }

        private void txt_hesoluong_KeyPress(object sender, KeyPressEventArgs e)
        {
            // ===== Cho nhập số =====
            if (!char.IsControl(e.KeyChar)
                && !char.IsDigit(e.KeyChar)
                && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            // ===== Chỉ cho nhập 1 dấu chấm =====
            if (e.KeyChar == '.'
                && txt_hesoluong.Text.Contains("."))
            {
                e.Handled = true;
            }

            // ===== Chỉ cho 2 số sau dấu chấm =====
            if (txt_hesoluong.Text.Contains("."))
            {
                string[] part =
                    txt_hesoluong.Text.Split('.');

                if (part.Length > 1)
                {
                    if (part[1].Length >= 2
                        && !char.IsControl(e.KeyChar))
                    {
                        e.Handled = true;
                    }
                }
            }
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

        private void btn_xuat_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (gv_hopdong.RowCount <= 0)
                {
                    MessageBox.Show(
                        "Không có dữ liệu để xuất");
                    return;
                }

                SaveFileDialog save =
                    new SaveFileDialog();

                save.Filter =
                    "Excel Workbook|*.xlsx";

                save.Title =
                    "Xuất Excel";

                save.FileName =
                    "DanhSachHopDong.xlsx";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                // ===== Mở kết nối =====
                ConnectData.taoketnoi();

                // ===== Khởi tạo Excel =====
                Microsoft.Office.Interop.Excel.Application app =
                    new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook wb =
                    app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet ws =
                    (Microsoft.Office.Interop.Excel.Worksheet)
                    wb.ActiveSheet;

                ws.Name = "HOPDONG";

                int cot =
                    gv_hopdong.Columns.Count;

                int dong =
                    gv_hopdong.RowCount;

                // ===== TIÊU ĐỀ =====
                Microsoft.Office.Interop.Excel.Range title =
                    ws.Range["A1", "L1"];

                title.Merge();

                title.Value =
                    "DANH SÁCH HỢP ĐỒNG NHÂN VIÊN";

                title.Font.Bold = true;

                title.Font.Size = 18;

                title.HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // ===== HEADER =====
                int colExcel = 1;

                int colPhongBan = 0;

                for (int i = 0; i < cot; i++)
                {
                    string tenCot =
                        gv_hopdong.Columns[i].FieldName;

                    // ===== Bỏ cột sửa xóa =====
                    if (tenCot == "Sua"
                        || tenCot == "Xoa")
                        continue;

                    ws.Cells[3, colExcel] =
                        gv_hopdong.Columns[i].Caption;

                    ws.Cells[3, colExcel].Font.Bold =
                        true;

                    ws.Cells[3, colExcel].Borders.LineStyle =
                        1;

                    ws.Cells[3, colExcel].Interior.Color =
                        System.Drawing.Color.LightGray;

                    // ===== Ép TEXT =====
                    ws.Columns[colExcel].NumberFormat =
                        "@";

                    // ===== Cột phòng ban =====
                    if (tenCot == "PhongBan_ID")
                    {
                        colPhongBan = colExcel;
                    }

                    colExcel++;
                }

                // ===== DỮ LIỆU =====
                for (int i = 0; i < dong; i++)
                {
                    colExcel = 1;

                    for (int j = 0; j < cot; j++)
                    {
                        string tenCot =
                            gv_hopdong.Columns[j].FieldName;

                        if (tenCot == "Sua"
                            || tenCot == "Xoa")
                            continue;

                        object value;

                        // ===== Hiển thị tên phòng ban =====
                        if (tenCot == "PhongBan_ID")
                        {
                            value =
                                gv_hopdong.GetRowCellDisplayText(
                                    i,
                                    gv_hopdong.Columns[j]);
                        }
                        else
                        {
                            value =
                                gv_hopdong.GetRowCellValue(
                                    i,
                                    gv_hopdong.Columns[j]);
                        }

                        // ===== Chỉ format cột ngày =====
                        if (value != null &&
                           (
                            tenCot == "NgayBD"
                            || tenCot == "NgaySinh"
                            || tenCot == "NgayCap"
                            || tenCot == "NgayCapCCCD"
                            || tenCot == "CREATEO_DATE"
                            || tenCot == "UPDATEO_DATE"
                            || tenCot == "DELETEO_DATE"
                           ))
                        {
                            DateTime ngay;

                            if (DateTime.TryParse(
                                value.ToString(),
                                out ngay))
                            {
                                ws.Cells[i + 4, colExcel] =
                                    "'" + ngay.ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                ws.Cells[i + 4, colExcel] =
                                    "'" + value?.ToString();
                            }
                        }
                        else
                        {
                            // ===== Giữ nguyên =====
                            ws.Cells[i + 4, colExcel] =
                                "'" + value?.ToString();
                        }

                        // ===== Border =====
                        ws.Cells[i + 4, colExcel]
                            .Borders.LineStyle = 1;

                        colExcel++;
                    }
                }

                // ===== Lấy danh sách phòng ban =====
                List<string> dsPhongBan =
                    new List<string>();

                string sql =
                    "SELECT TenPhongBan FROM PHONGBAN";

                MySqlCommand cmd =
                    new MySqlCommand(
                        sql,
                        ConnectData.conn);

                MySqlDataReader dr =
                    cmd.ExecuteReader();

                while (dr.Read())
                {
                    dsPhongBan.Add(
                        dr["TenPhongBan"].ToString());
                }

                dr.Close();

                // ===== Sheet ẩn =====
                Microsoft.Office.Interop.Excel.Worksheet wsPB =
                    (Microsoft.Office.Interop.Excel.Worksheet)
                    wb.Sheets.Add();

                wsPB.Name =
                    "PHONGBAN_DATA";

                for (int i = 0; i < dsPhongBan.Count; i++)
                {
                    wsPB.Cells[i + 1, 1] =
                        dsPhongBan[i];
                }

                // ===== Ẩn sheet =====
                wsPB.Visible =
                    Microsoft.Office.Interop.Excel
                    .XlSheetVisibility.xlSheetHidden;

                // ===== Dropdown combobox =====
                Microsoft.Office.Interop.Excel.Range rng =
                    ws.Range[
                        ws.Cells[4, colPhongBan],
                        ws.Cells[dong + 100, colPhongBan]
                    ];

                rng.Validation.Delete();

                rng.Validation.Add(
                    Microsoft.Office.Interop.Excel
                    .XlDVType.xlValidateList,

                    Microsoft.Office.Interop.Excel
                    .XlDVAlertStyle.xlValidAlertStop,

                    Microsoft.Office.Interop.Excel
                    .XlFormatConditionOperator.xlBetween,

                    "=PHONGBAN_DATA!$A$1:$A$"
                    + dsPhongBan.Count,

                    Type.Missing);

                rng.Validation.IgnoreBlank =
                    true;

                rng.Validation.InCellDropdown =
                    true;

                // ===== Căn giữa =====
                ws.Range["A3",
                    ws.Cells[dong + 3, colExcel - 1]]
                    .HorizontalAlignment =
                    Microsoft.Office.Interop.Excel
                    .XlHAlign.xlHAlignCenter;

                // ===== AutoFit =====
                ws.Columns.AutoFit();

                // ===== Lưu =====
                wb.SaveAs(save.FileName);

                wb.Close();

                app.Quit();

                // ===== Giải phóng =====
                System.Runtime.InteropServices
                    .Marshal.ReleaseComObject(ws);

                System.Runtime.InteropServices
                    .Marshal.ReleaseComObject(wsPB);

                System.Runtime.InteropServices
                    .Marshal.ReleaseComObject(wb);

                System.Runtime.InteropServices
                    .Marshal.ReleaseComObject(app);

                MessageBox.Show(
                    "Xuất Excel thành công",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // ===== Mở file =====
                System.Diagnostics.Process.Start(
                    save.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ConnectData.dongketnoi();
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
                    try
                    {
                        // ===== Mã nhân viên =====
                        string maNV =
                            Convert.ToString(
                                (range.Cells[i, 2] as Excel.Range).Text);

                        if (string.IsNullOrWhiteSpace(maNV))
                            continue;

                        // ===== Lấy ID nhân viên =====
                        string sqlNhanVien =
                            "SELECT id FROM NHANVIEN WHERE MaNV=@MaNV AND DELETEO_BY IS NULL";

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
                                "Không tìm thấy nhân viên mã: "
                                + maNV);

                            continue;
                        }

                        int nhanVienID =
                            Convert.ToInt32(result);

                        // ===== Check hợp đồng =====
                        string sqlCheckNV =
                            "SELECT COUNT(*) FROM HOPDONG WHERE NhanVien_id=@NhanVien_id";

                        MySqlCommand cmdCheckNV =
                            new MySqlCommand(
                                sqlCheckNV,
                                ConnectData.conn);

                        cmdCheckNV.Parameters.AddWithValue(
                            "@NhanVien_id",
                            nhanVienID);

                        int checkNV =
                            Convert.ToInt32(
                                cmdCheckNV.ExecuteScalar());

                        if (checkNV > 0)
                        {
                            MessageBox.Show(
                                "Nhân viên "
                                + maNV
                                + " đã có hợp đồng");

                            continue;
                        }

                        // ===== Số hợp đồng =====
                        string soHopDong =
                            Convert.ToString(
                                (range.Cells[i, 7] as Excel.Range).Text);

                        if (string.IsNullOrWhiteSpace(
                            soHopDong))
                        {
                            MessageBox.Show(
                                "Dòng "
                                + i
                                + " chưa có số hợp đồng");

                            continue;
                        }

                        // ===== Check trùng số hợp đồng =====
                        string sqlCheck =
                            "SELECT COUNT(*) FROM HOPDONG WHERE So_HopDong=@So_HopDong";

                        MySqlCommand cmdCheck =
                            new MySqlCommand(
                                sqlCheck,
                                ConnectData.conn);

                        cmdCheck.Parameters.AddWithValue(
                            "@So_HopDong",
                            soHopDong);

                        int check =
                            Convert.ToInt32(
                                cmdCheck.ExecuteScalar());

                        if (check > 0)
                        {
                            MessageBox.Show(
                                "Số hợp đồng "
                                + soHopDong
                                + " đã tồn tại");

                            continue;
                        }

                        // ===== Ngày bắt đầu =====
                        DateTime ngayBD;

                        Excel.Range cellNgayBD =
                            (Excel.Range)range.Cells[i, 8];

                        double oaDate = 0;

                        if (cellNgayBD.Value2 != null
                            && double.TryParse(
                                cellNgayBD.Value2.ToString(),
                                out oaDate))
                        {
                            ngayBD =
                                DateTime.FromOADate(
                                    oaDate);
                        }
                        else if (!DateTime.TryParse(
                            cellNgayBD.Text.ToString(),
                            out ngayBD))
                        {
                            MessageBox.Show(
                                "Lỗi ngày bắt đầu dòng "
                                + i);

                            continue;
                        }

                        // ===== Loại hợp đồng =====
                        string loaiHopDong =
                            Convert.ToString(
                                (range.Cells[i, 9] as Excel.Range).Text);

                        // ===== Phòng ban =====
                        string tenPhongBan =
                            Convert.ToString(
                                (range.Cells[i, 10] as Excel.Range).Text);

                        if (string.IsNullOrWhiteSpace(
                            tenPhongBan))
                        {
                            MessageBox.Show(
                                "Dòng "
                                + i
                                + " chưa có phòng ban");

                            continue;
                        }

                        string sqlPhongBan =
                            "SELECT id FROM PHONGBAN WHERE TenPhongBan=@TenPhongBan";

                        MySqlCommand cmdPhongBan =
                            new MySqlCommand(
                                sqlPhongBan,
                                ConnectData.conn);

                        cmdPhongBan.Parameters.AddWithValue(
                            "@TenPhongBan",
                            tenPhongBan);

                        object resultPhongBan =
                            cmdPhongBan.ExecuteScalar();

                        if (resultPhongBan == null)
                        {
                            MessageBox.Show(
                                "Không tìm thấy phòng ban: "
                                + tenPhongBan);

                            continue;
                        }

                        int phongBanID =
                            Convert.ToInt32(
                                resultPhongBan);

                        // ===== Chức vụ =====
                        string chucVu =
                            Convert.ToString(
                                (range.Cells[i, 11] as Excel.Range).Text);

                        // ===== Lương cơ bản =====
                        int luongCB = 0;

                        Excel.Range cellLuongCB =
                            (Excel.Range)range.Cells[i, 12];

                        string luongCBText =
                            cellLuongCB.Text.ToString()
                            .Replace(".", "")
                            .Replace(",", "")
                            .Trim();

                        if (!int.TryParse(
                            luongCBText,
                            out luongCB))
                        {
                            MessageBox.Show(
                                "Lỗi lương cơ bản dòng "
                                + i);

                            continue;
                        }

                        // ===== Hệ số lương =====
                        float heSoLuong = 0;

                        Excel.Range cellHeSoLuong =
                            (Excel.Range)range.Cells[i, 13];

                        string heSoLuongText =
                            cellHeSoLuong.Text.ToString()
                            .Replace(",", ".")
                            .Trim();

                        if (!float.TryParse(
                            heSoLuongText,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out heSoLuong))
                        {
                            MessageBox.Show(
                                "Lỗi hệ số lương dòng "
                                + i);

                            continue;
                        }

                        // ===== Ghi chú =====
                        string ghiChu =
                            Convert.ToString(
                                (range.Cells[i, 14] as Excel.Range).Text);

                        // ===== Insert =====
                        string sql = @"
INSERT INTO HOPDONG
(
    NhanVien_id,
    So_HopDong,
    NgayBD,
    LoaiHopDong,
    PhongBan_ID,
    ChucVu,
    LuongCB,
    HeSoLuong,
    GhiChu
)
VALUES
(
    @NhanVien_id,
    @So_HopDong,
    @NgayBD,
    @LoaiHopDong,
    @PhongBan_ID,
    @ChucVu,
    @LuongCB,
    @HeSoLuong,
    @GhiChu
)";

                        MySqlCommand cmd =
                            new MySqlCommand(
                                sql,
                                ConnectData.conn);

                        cmd.Parameters.AddWithValue(
                            "@NhanVien_id",
                            nhanVienID);

                        cmd.Parameters.AddWithValue(
                            "@So_HopDong",
                            soHopDong);

                        cmd.Parameters.AddWithValue(
                            "@NgayBD",
                            ngayBD);

                        cmd.Parameters.AddWithValue(
                            "@LoaiHopDong",
                            loaiHopDong);

                        cmd.Parameters.AddWithValue(
                            "@PhongBan_ID",
                            phongBanID);

                        cmd.Parameters.AddWithValue(
                            "@ChucVu",
                            chucVu);

                        cmd.Parameters.AddWithValue(
                            "@LuongCB",
                            luongCB);

                        cmd.Parameters.AddWithValue(
                            "@HeSoLuong",
                            heSoLuong);

                        cmd.Parameters.AddWithValue(
                            "@GhiChu",
                            ghiChu);

                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception exRow)
                    {
                        MessageBox.Show(
                            "Lỗi dòng "
                            + i
                            + "\n"
                            + exRow.Message);
                    }
                }

                wb.Close(false);

                app.Quit();

                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);

                MessageBox.Show(
                    "Import hợp đồng thành công",
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

            load_hopdong();
            Loadnhanvien();
        }

        private void lbl_phongban_EditValueChanged(object sender, EventArgs e)
        {

        }
    }
}