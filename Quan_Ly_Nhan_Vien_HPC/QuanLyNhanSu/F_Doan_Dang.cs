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
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class F_Doan_Dang : DevExpress.XtraEditors.XtraForm
    {
        public F_Doan_Dang()
        {
            InitializeComponent();
        }

        void Loadnhanvien()
        {
            this.lku_nhanvien.Properties.DataSource = ConnectData.getdata("SELECT \r\n        nv.*,\r\n        CONCAT(nv.HoDem, ' ', nv.Ten) AS HoTen\r\n    FROM NHANVIEN nv\r\n    LEFT JOIN THONGTINDANG bh\r\n        ON nv.id = bh.NhanVien_id\r\n    WHERE nv.DELETEO_BY IS NULL\r\n        AND bh.NhanVien_id IS NULL\r\n    ORDER BY nv.id ASC");
            this.lku_nhanvien.Properties.DisplayMember = "HoTen";
            this.lku_nhanvien.Properties.ValueMember = "id";

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

                ConnectData.taoketnoi();

                int nhanVienID =
                    Convert.ToInt32(
                        lku_nhanvien.EditValue);

                // ===== Check mỗi nhân viên chỉ có 1 thông tin đảng =====
                string sqlCheck =
                    @"SELECT COUNT(*)
              FROM THONGTINDANG
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
                        "Nhân viên này đã có thông tin đảng");

                    return;
                }

                // ===== Ngày vào đoàn =====
                DateTime ngayVaoDoan =
                    date_NgayVaoDoan.Value.Date;

                // ===== Ngày vào đảng =====
                DateTime? ngayVaoDang = null;

                if (!string.IsNullOrWhiteSpace(
                    date_ngayvaodang.Text))
                {
                    DateTime tempNgayDang;

                    bool checkNgayDang =
                        DateTime.TryParseExact(
                            date_ngayvaodang.Text.Trim(),
                            "dd/MM/yyyy",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out tempNgayDang);

                    if (!checkNgayDang)
                    {
                        MessageBox.Show(
                            "Ngày vào đảng không đúng định dạng dd/MM/yyyy");

                        date_ngayvaodang.Focus();
                        return;
                    }

                    ngayVaoDang =
                        tempNgayDang;
                }

                // ===== Ngày chính thức vào đảng =====
                DateTime? ngayChinhThuc = null;

                if (!string.IsNullOrWhiteSpace(
                    date_ngaychinhthucvaodang.Text))
                {
                    DateTime tempNgayCT;

                    bool checkNgayCT =
                        DateTime.TryParseExact(
                            date_ngaychinhthucvaodang.Text.Trim(),
                            "dd/MM/yyyy",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out tempNgayCT);

                    if (!checkNgayCT)
                    {
                        MessageBox.Show(
                            "Ngày chính thức vào đảng không đúng định dạng dd/MM/yyyy");

                        date_ngaychinhthucvaodang.Focus();
                        return;
                    }

                    ngayChinhThuc =
                        tempNgayCT;
                }

                // ===== INSERT =====
                string sql = @"
                    INSERT INTO THONGTINDANG
                    (
                        NhanVien_id,
                        NGAYVAODOAN,
                        NGAYVAODANG,
                        NgayChinhThucVaoDang,
                        GhiChu
                    )
                    VALUES
                    (
                        @NhanVien_id,
                        @NGAYVAODOAN,
                        @NGAYVAODANG,
                        @NgayChinhThucVaoDang,
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
                    "@NGAYVAODOAN",
                    ngayVaoDoan);

                // ===== Ngày vào đảng =====
                if (ngayVaoDang.HasValue)
                {
                    cmd.Parameters.AddWithValue(
                        "@NGAYVAODANG",
                        ngayVaoDang.Value.ToString("yyyy-MM-dd"));
                }
                else
                {
                    cmd.Parameters.AddWithValue(
                        "@NGAYVAODANG",
                        DBNull.Value);
                }

                // ===== Ngày chính thức =====
                if (ngayChinhThuc.HasValue)
                {
                    cmd.Parameters.AddWithValue(
                        "@NgayChinhThucVaoDang",
                        ngayChinhThuc.Value.ToString("yyyy-MM-dd"));
                }
                else
                {
                    cmd.Parameters.AddWithValue(
                        "@NgayChinhThucVaoDang",
                        DBNull.Value);
                }

                cmd.Parameters.AddWithValue(
                    "@GhiChu",
                    txt_ghichu.Text.Trim());

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Thêm thông tin đảng thành công",
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

            load_dang();
            Loadnhanvien();
        }
        private void load_dang()
        {
            gc_dang.DataSource = ConnectData.getdata("SELECT \r\n    td.Id,\r\n    nv.MaNV,\r\n    CONCAT(nv.HoDem, ' ', nv.Ten) AS HoTen,\r\n    nv.NgaySinh,\r\n    nv.SDT,\r\n    nv.Email,\r\n    td.NGAYVAODOAN,\r\n    td.NGAYVAODANG,\r\n    td.NgayChinhThucVaoDang,\r\n    td.GhiChu\r\nFROM THONGTINDANG td\r\nINNER JOIN NHANVIEN nv\r\n    ON td.NhanVien_id = nv.id\r\nWHERE nv.DELETEO_BY IS NULL\r\nORDER BY td.Id DESC;\r\n");
        }
        private void F_Doan_Dang_Load(object sender, EventArgs e)
        {


            lku_nhanvien.Focus();
            load_dang();

            Loadnhanvien();

            gv_dang.OptionsBehavior.Editable = true;
            gv_dang.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None;
            GridView gridView = gv_dang;
            // Định dạng căn giữa chữ trong Header
            gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            // Định dạng căn giữa chữ trong Row
            gridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            btn_sua.Click += Btn_sua_Click;
            btn_xoa.Click += Btn_xoa_Click;
            lku_nhanvien.KeyDown += lku_nhanvien_KeyDown;
            date_NgayVaoDoan.KeyDown += date_NgayVaoDoan_KeyDown;
            date_ngayvaodang.KeyDown += date_ngayvaodang_KeyDown;
            date_ngaychinhthucvaodang.KeyDown += date_ngaychinhthucvaodang_KeyDown;
            txt_ghichu.KeyDown += txt_ghichu_KeyDown;
            //txt_luongcb.KeyDown += txt_luongcb_KeyDown;
            //txt_ghichu.KeyDown += txt_ghichu_KeyDown;
            //txt_hesoluong.KeyDown += txt_hesoluong_KeyDown;
            //cb_loaihopdong.KeyDown += cb_loaihopdong_KeyDown;
            //date_NgayBD.KeyDown += date_NgayBD_KeyDown;
            //txt_chucvu.KeyDown += txt_chucvu_KeyDown;
            //lbl_phongban.KeyDown += lbl_phongban_KeyDown;
            if (gv_dang.Columns["STT"] != null)
                gv_dang.Columns["STT"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_dang.Columns["MaNV"] != null)
                gv_dang.Columns["MaNV"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_dang.Columns["HoTen"] != null)
                gv_dang.Columns["HoTen"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_dang.Columns["btn_xoa"] != null)
                gv_dang.Columns["btn_xoa"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;

            if (gv_dang.Columns["btn_sua"] != null)
                gv_dang.Columns["btn_sua"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;



            // ===== Tự động độ rộng =====
            gv_dang.OptionsView.ColumnAutoWidth = false;
            gv_dang.BestFitColumns();

            // ===== Thanh cuộn =====
            gv_dang.HorzScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;

            gv_dang.VertScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            gv_dang.Columns["HoTen"].Width = 200;

            date_ngayvaodang.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTime;
            date_ngayvaodang.Properties.Mask.EditMask = "dd/MM/yyyy";
            date_ngayvaodang.Properties.Mask.UseMaskAsDisplayFormat = true;

            date_ngaychinhthucvaodang.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTime;
            date_ngaychinhthucvaodang.Properties.Mask.EditMask = "dd/MM/yyyy";
            date_ngaychinhthucvaodang.Properties.Mask.UseMaskAsDisplayFormat = true;
        }

        private void txt_ghichu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_them.PerformClick();
            }
        }

        private void date_ngaychinhthucvaodang_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_ghichu.Focus();
            }
        }

        private void date_ngayvaodang_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                date_ngaychinhthucvaodang.Focus();
            }
        }

        private void date_NgayVaoDoan_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                date_ngayvaodang.Focus();
            }
        }

        private void lku_nhanvien_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                date_NgayVaoDoan.Focus();
            }
        }

        private void Btn_xoa_Click(object sender, EventArgs e)
        {
            try
            {
                // ===== Check chọn dòng =====
                if (gv_dang.FocusedRowHandle < 0)
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
                        gv_dang.GetFocusedRowCellValue("Id"));

                // ===== Lấy họ tên =====
                string hoten =
                    Convert.ToString(
                        gv_dang.GetFocusedRowCellValue("HoTen"));

                // ===== Xác nhận =====
                DialogResult rs =
                    MessageBox.Show(
                        "Bạn có chắc chắn muốn xóa thông tin đảng của nhân viên: "
                        + hoten + " ?",
                        "Thông báo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                if (rs != DialogResult.Yes)
                    return;

                ConnectData.taoketnoi();

                // ===== SQL DELETE =====
                string sql =
                    "DELETE FROM THONGTINDANG WHERE Id=@Id";

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
                        "Xóa thông tin đảng thành công",
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

            // ===== Reload =====
            load_dang();
            Loadnhanvien();
        }

        private void Btn_sua_Click(object sender, EventArgs e)
        {
            try
            {
                if (gv_dang.FocusedRowHandle < 0)
                    return;

                string hoTen =
               gv_dang.GetFocusedRowCellValue("HoTen")?.ToString();

                DialogResult rs = MessageBox.Show(
                    "Bạn có chắc chắn muốn sửa thông tin đảng của:\n"
                    + hoTen + " không?",
                    "Thông báo",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (rs == DialogResult.No)
                    return;

                // ===== ID =====
                int id =
                    Convert.ToInt32(
                        gv_dang.GetFocusedRowCellValue("Id"));

                // ===== Ngày vào đoàn =====
                DateTime ngayVaoDoan =
                    Convert.ToDateTime(
                        gv_dang.GetFocusedRowCellValue("NGAYVAODOAN"));

                // ===== Ngày vào đảng =====
                object objNgayDang =
                    gv_dang.GetFocusedRowCellValue("NGAYVAODANG");

                DateTime? ngayVaoDang = null;

                if (objNgayDang != null
                    && objNgayDang != DBNull.Value
                    && !string.IsNullOrWhiteSpace(
                        objNgayDang.ToString()))
                {
                    ngayVaoDang =
                        Convert.ToDateTime(objNgayDang);
                }

                // ===== Ngày chính thức =====
                object objNgayCT =
                    gv_dang.GetFocusedRowCellValue(
                        "NgayChinhThucVaoDang");

                DateTime? ngayChinhThuc = null;

                if (objNgayCT != null
                    && objNgayCT != DBNull.Value
                    && !string.IsNullOrWhiteSpace(
                        objNgayCT.ToString()))
                {
                    ngayChinhThuc =
                        Convert.ToDateTime(objNgayCT);
                }

                // ===== Ghi chú =====
                string ghiChu =
                    Convert.ToString(
                        gv_dang.GetFocusedRowCellValue("GhiChu"));

                ConnectData.taoketnoi();

                // ===== UPDATE =====
                string sql = @"
                    UPDATE THONGTINDANG
                    SET
                        NGAYVAODOAN=@NGAYVAODOAN,
                        NGAYVAODANG=@NGAYVAODANG,
                        NgayChinhThucVaoDang=@NgayChinhThucVaoDang,
                        GhiChu=@GhiChu
                    WHERE Id=@Id";

                MySqlCommand cmd =
                    new MySqlCommand(
                        sql,
                        ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@Id",
                    id);

                cmd.Parameters.AddWithValue(
                    "@NGAYVAODOAN",
                    ngayVaoDoan);

                // ===== Ngày vào đảng =====
                if (ngayVaoDang.HasValue)
                {
                    cmd.Parameters.AddWithValue(
                        "@NGAYVAODANG",
                        ngayVaoDang.Value.ToString("yyyy-MM-dd"));
                }
                else
                {
                    cmd.Parameters.AddWithValue(
                        "@NGAYVAODANG",
                        DBNull.Value);
                }

                // ===== Ngày chính thức =====
                if (ngayChinhThuc.HasValue)
                {
                    cmd.Parameters.AddWithValue(
                        "@NgayChinhThucVaoDang",
                        ngayChinhThuc.Value.ToString("yyyy-MM-dd"));
                }
                else
                {
                    cmd.Parameters.AddWithValue(
                        "@NgayChinhThucVaoDang",
                        DBNull.Value);
                }

                cmd.Parameters.AddWithValue(
                    "@GhiChu",
                    ghiChu);

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Sửa thông tin đảng thành công",
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

            load_dang();
            Loadnhanvien();
            clearInput();
        }
        void clearInput()
        {
            lku_nhanvien.EditValue = null;
            
            date_ngayvaodang.EditValue = null;
            date_ngaychinhthucvaodang.EditValue = null;
            txt_ghichu.Text = string.Empty;
        }
        private void btn_lammoi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            clearInput();
            load_dang();
             Loadnhanvien();
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
                    "DanhSachThongTinDang.xlsx";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                Microsoft.Office.Interop.Excel.Application app =
                    new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook wb =
                    app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet ws =
                    (Microsoft.Office.Interop.Excel.Worksheet)
                    wb.ActiveSheet;

                ws.Name = "THONGTINDANG";

                int cot =
                    gv_dang.Columns.Count;

                int dong =
                    gv_dang.RowCount;

                // ===== TIÊU ĐỀ =====
                Microsoft.Office.Interop.Excel.Range title =
                    ws.Range["A1", "K1"];

                title.Merge();

                title.Value =
                    "DANH SÁCH THÔNG TIN ĐOÀN VÀ ĐẢNG";

                title.Font.Bold = true;

                title.Font.Size = 18;

                title.HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // ===== HEADER =====
                int colExcel = 1;

                for (int i = 0; i < cot; i++)
                {
                    string tenCot =
                        gv_dang.Columns[i].FieldName;

                    // ===== Bỏ cột ID =====
                    if (tenCot == "Id")
                        continue;

                    ws.Cells[3, colExcel] =
                        gv_dang.Columns[i].Caption;

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
                            gv_dang.Columns[j].FieldName;

                        if (tenCot == "Id")
                            continue;

                        object value =
                            gv_dang.GetRowCellValue(
                                i,
                                gv_dang.Columns[j]);

                        // ===== Format ngày =====
                        if (value != null &&
                           (
                            tenCot == "NgaySinh"
                            || tenCot == "NGAYVAODOAN"
                            || tenCot == "NGAYVAODANG"
                            || tenCot == "NgayChinhThucVaoDang"
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
                  FROM THONGTINDANG
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
                            + " đã có thông tin đảng");

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
INSERT INTO THONGTINDANG
(
    NhanVien_id,
    NGAYVAODOAN,
    NGAYVAODANG,
    NgayChinhThucVaoDang,
    GhiChu
)
VALUES
(
    @NhanVien_id,
    @NGAYVAODOAN,
    @NGAYVAODANG,
    @NgayChinhThucVaoDang,
    @GhiChu
)";

                    MySqlCommand cmd =
                        new MySqlCommand(
                            sql,
                            ConnectData.conn);

                    cmd.Parameters.AddWithValue(
                        "@NhanVien_id",
                        nhanVienID);

                    // ===== Ngày vào đoàn =====
                    cmd.Parameters.AddWithValue(
                        "@NGAYVAODOAN",
                        ConvertDate(7));

                    // ===== Ngày vào đảng =====
                    cmd.Parameters.AddWithValue(
                        "@NGAYVAODANG",
                        ConvertDate(8));

                    // ===== Ngày chính thức =====
                    cmd.Parameters.AddWithValue(
                        "@NgayChinhThucVaoDang",
                        ConvertDate(9));

                    // ===== Ghi chú =====
                    cmd.Parameters.AddWithValue(
                        "@GhiChu",
                        GetText(10));

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

            load_dang();
            Loadnhanvien();
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