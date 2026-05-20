using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class F_BaoHiem : DevExpress.XtraEditors.XtraForm
    {
        public F_BaoHiem()
        {
            InitializeComponent();
        }
        void Loadnhanvien()
        {
            this.lku_nhanvien.Properties.DataSource = ConnectData.getdata("SELECT \r\n        nv.*,\r\n        CONCAT(nv.HoDem, ' ', nv.Ten) AS HoTen\r\n    FROM NHANVIEN nv\r\n    LEFT JOIN BAOHIEM bh\r\n        ON nv.id = bh.NhanVien_id\r\n    WHERE nv.DELETEO_BY IS NULL\r\n        AND bh.NhanVien_id IS NULL\r\n    ORDER BY nv.id ASC");
            this.lku_nhanvien.Properties.DisplayMember = "HoTen";
            this.lku_nhanvien.Properties.ValueMember = "id";
        }
        private void F_BaoHiem_Load(object sender, EventArgs e)
        {
            Loadnhanvien();
            load_baohiem();
            gv_phonban.OptionsBehavior.Editable = true;
            gv_phonban.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None;
            GridView gridView = gv_phonban;
            // Định dạng căn giữa chữ trong Header
            gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            // Định dạng căn giữa chữ trong Row
            gridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            btn_sua.Click += Btn_sua_Click;
            btn_xoa.Click += Btn_xoa_Click;
            lku_nhanvien.KeyDown += lku_nhanvien_KeyDown;
            txt_sobaohiem.KeyDown += txt_sobaohiem_KeyDown;
            txt_ghichu.KeyDown += txt_ghichu_KeyDown;
            txt_noikham.KeyDown += txt_noikham_KeyDown;
            txt_noicap.KeyDown += txt_noicap_KeyDown;
            date_cap.KeyDown += date_cap_KeyDown;
            if (gv_phonban.Columns["STT"] != null)
                gv_phonban.Columns["STT"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_phonban.Columns["MaNV"] != null)
                gv_phonban.Columns["MaNV"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_phonban.Columns["HoTen"] != null)
                gv_phonban.Columns["HoTen"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Left;

            if (gv_phonban.Columns["btn_xoa"] != null)
                gv_phonban.Columns["btn_xoa"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;

            if (gv_phonban.Columns["btn_sua"] != null)
                gv_phonban.Columns["btn_sua"].Fixed =
                    DevExpress.XtraGrid.Columns.FixedStyle.Right;



            // ===== Tự động độ rộng =====
            gv_phonban.OptionsView.ColumnAutoWidth = false;
            gv_phonban.BestFitColumns();

            // ===== Thanh cuộn =====
            gv_phonban.HorzScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;

            gv_phonban.VertScrollVisibility =
                DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
        }

        private void date_cap_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_ghichu.Focus();
            }
        }

        private void txt_noicap_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                date_cap.Focus();
            }
        }

        private void txt_noikham_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_noicap.Focus();
            }
        }

        private void txt_ghichu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_themBH.PerformClick();
            }
        }

        private void txt_sobaohiem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_noikham.Focus();
            }
        }

        private void lku_nhanvien_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_sobaohiem.Focus();
            }
        }

        private void Btn_xoa_Click(object sender, EventArgs e)
        {
            if (gv_phonban.FocusedRowHandle < 0)
            {
                MessageBox.Show("Chọn dữ liệu cần xóa");
                return;
            }

            string hoTen =
                gv_phonban.GetFocusedRowCellValue("HoTen")?.ToString();

            string soBaoHiem =
                gv_phonban.GetFocusedRowCellValue("So_BaoHiem")?.ToString();

            DialogResult rs = MessageBox.Show(
                "Bạn có chắc chắn muốn xóa bảo hiểm của:\n"
                + hoTen
                + "\n\nSố bảo hiểm: "
                + soBaoHiem
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
                        gv_phonban.GetFocusedRowCellValue("Id"));

                string sql =
                    "DELETE FROM BAOHIEM WHERE Id=@Id";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@Id",
                    id);

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Xóa bảo hiểm thành công của:\n"
                    + hoTen,
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                
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
            load_baohiem();
            Loadnhanvien();
        }

        private void Btn_sua_Click(object sender, EventArgs e)
        {
            if (gv_phonban.FocusedRowHandle < 0)
            {
                MessageBox.Show("Chọn dữ liệu cần sửa");
                return;
            }

            string hoTen =
                gv_phonban.GetFocusedRowCellValue("HoTen")?.ToString();

            DialogResult rs = MessageBox.Show(
                "Bạn có chắc chắn muốn sửa bảo hiểm của:\n"
                + hoTen + " không?",
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
                        gv_phonban.GetFocusedRowCellValue("Id"));

                string soBaoHiem =
                    gv_phonban.GetFocusedRowCellValue("So_BaoHiem")?.ToString();

                DateTime ngayCap =
                    Convert.ToDateTime(
                        gv_phonban.GetFocusedRowCellValue("NgayCap"));
                string noiCap =
                    gv_phonban.GetFocusedRowCellValue("NoiCap")?.ToString();

                string noiDK =
                    gv_phonban.GetFocusedRowCellValue("Noi_DK_Kham_Benh")?.ToString();

                string ghiChu =
                    gv_phonban.GetFocusedRowCellValue("GhiChu")?.ToString();

                // ===== Check rỗng =====
                if (string.IsNullOrWhiteSpace(soBaoHiem))
                {
                    MessageBox.Show("Số bảo hiểm không được để trống");
                    return;
                }

                // ===== Check trùng số bảo hiểm =====
                string sqlCheck = @"
        SELECT COUNT(*)
        FROM BAOHIEM
        WHERE So_BaoHiem=@So_BaoHiem
        AND Id<>@Id";

                MySqlCommand cmdCheck =
                    new MySqlCommand(sqlCheck, ConnectData.conn);

                cmdCheck.Parameters.AddWithValue(
                    "@So_BaoHiem",
                    soBaoHiem);

                cmdCheck.Parameters.AddWithValue(
                    "@Id",
                    id);

                int check =
                    Convert.ToInt32(
                        cmdCheck.ExecuteScalar());

                if (check > 0)
                {
                    MessageBox.Show(
                        "Số bảo hiểm đã tồn tại");

                    return;
                }

                // ===== Update =====
                string sql = @"
        UPDATE BAOHIEM
        SET
            So_BaoHiem=@So_BaoHiem,
            NgayCap=@NgayCap,
            NoiCap=@NoiCap,
            Noi_DK_Kham_Benh=@NoiDK,
            GhiChu=@GhiChu
        WHERE Id=@Id";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@So_BaoHiem",
                    soBaoHiem);

                cmd.Parameters.AddWithValue(
                    "@NgayCap",
                    ngayCap);

                cmd.Parameters.AddWithValue(
                    "@NoiCap",
                    noiCap);

                cmd.Parameters.AddWithValue(
                    "@NoiDK",
                    noiDK);

                cmd.Parameters.AddWithValue(
                    "@GhiChu",
                    ghiChu);

                cmd.Parameters.AddWithValue(
                    "@Id",
                    id);

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Sửa bảo hiểm thành công cho:\n"
                    + hoTen,
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
            load_baohiem();
        }

        private void btn_them_Click(object sender, EventArgs e)
        {
        }

        private void load_baohiem()
        {
            gc_phongban.DataSource = ConnectData.getdata("SELECT bh.Id, nv.MaNV, CONCAT(nv.HoDem,' ',nv.Ten) AS HoTen,nv.NgaySinh,nv.SDT,nv.Email, bh.So_BaoHiem, bh.NgayCap, bh.NoiCap, bh.Noi_DK_Kham_Benh, bh.GhiChu\r\nFROM BAOHIEM bh\r\nINNER JOIN NHANVIEN nv\r\n    ON bh.NhanVien_id = nv.id\r\nWHERE nv.DELETEO_BY IS NULL\r\nORDER BY bh.Id desc;");
        }
        void clearform()
        {
            lku_nhanvien.EditValue = null;
            txt_sobaohiem.Text = "";
            date_cap.Value = DateTime.Now;
            txt_noicap.Text = "";
            txt_noikham.Text = "";
            txt_ghichu.Text = "";
        }

        private void btn_lammoi_Click(object sender, EventArgs e)
        {
            
        }

        private void btn_xuatexcel_Click(object sender, EventArgs e)
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
                    "DanhSachThongTinBaoHiem.xlsx";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                Excel.Application app =
                    new Excel.Application();

                Excel.Workbook wb =
                    app.Workbooks.Add(Type.Missing);

                Excel.Worksheet ws =
                    (Excel.Worksheet)wb.ActiveSheet;

                ws.Name = "BaoHiem";

                int rowCount = gv_phonban.RowCount;

                int totalCol = 0;

                // ===== Đếm cột =====
                for (int i = 0; i < gv_phonban.Columns.Count; i++)
                {
                    string colName =
                        gv_phonban.Columns[i].FieldName;

                    if (colName == "Sua"
                        || colName == "Xoa")
                        continue;

                    totalCol++;
                }

                // ===== Tiêu đề =====
                ws.Range[
                    ws.Cells[1, 1],
                    ws.Cells[1, totalCol]
                ].Merge();

                ws.Cells[1, 1] =
                    "DANH SÁCH BẢO HIỂM";

                Excel.Range title =
                    ws.Range[
                        ws.Cells[1, 1],
                        ws.Cells[1, totalCol]
                    ];

                title.Font.Bold = true;
                title.Font.Size = 18;

                title.HorizontalAlignment =
                    Excel.XlHAlign.xlHAlignCenter;

                // ===== Header =====
                int excelCol = 1;

                for (int i = 0; i < gv_phonban.Columns.Count; i++)
                {
                    string colName =
                        gv_phonban.Columns[i].FieldName;

                    if (colName == "Sua"
                        || colName == "Xoa")
                        continue;

                    ws.Cells[3, excelCol] =
                        gv_phonban.Columns[i].Caption;

                    excelCol++;
                }

                Excel.Range header =
                    ws.Range[
                        ws.Cells[3, 1],
                        ws.Cells[3, totalCol]
                    ];

                header.Font.Bold = true;

                // ===== Data =====
                for (int i = 0; i < rowCount; i++)
                {
                    excelCol = 1;

                    for (int j = 0; j < gv_phonban.Columns.Count; j++)
                    {
                        string colName =
                            gv_phonban.Columns[j].FieldName;

                        if (colName == "Sua"
                            || colName == "Xoa")
                            continue;

                        object value =
                            gv_phonban.GetRowCellValue(
                                i,
                                gv_phonban.Columns[j]);

                        // ===== Ép text =====
                        if (colName == "So_BaoHiem"
                            || colName == "SDT")
                        {
                            ws.Cells[i + 4, excelCol] =
                                "'" + value?.ToString();
                        }
                        else
                        {
                            ws.Cells[i + 4, excelCol] =
                                value;
                        }

                        excelCol++;
                    }
                }

                // ===== Border =====
                Excel.Range border =
                    ws.Range[
                        ws.Cells[3, 1],
                        ws.Cells[rowCount + 3, totalCol]
                    ];

                border.Borders.LineStyle =
                    Excel.XlLineStyle.xlContinuous;

                border.Borders.Weight =
                    Excel.XlBorderWeight.xlThin;

                // ===== AutoFit =====
                ws.Columns.AutoFit();

                wb.SaveAs(save.FileName);

                wb.Close();
                app.Quit();

                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);

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

        private void buttonEdit1_EditValueChanged(object sender, EventArgs e)
        {
            
        }

        
        private void btn_refest_EditValueChanged(object sender, EventArgs e)
        {
            
        }

        private void btn_importdulieu_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open = new OpenFileDialog();

                open.Filter = "Excel File|*.xlsx;*.xls";

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

                int rowCount = range.Rows.Count;

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

                    // ===== Lấy ID nhân viên =====
                    string sqlNhanVien =
                        "SELECT id FROM NHANVIEN WHERE MaNV=@MaNV AND DELETEO_BY IS NULL";

                    MySqlCommand cmdNhanVien =
                        new MySqlCommand(sqlNhanVien, ConnectData.conn);

                    cmdNhanVien.Parameters.AddWithValue(
                        "@MaNV",
                        maNV);

                    object result =
                        cmdNhanVien.ExecuteScalar();

                    if (result == null)
                    {
                        MessageBox.Show(
                            "Không tìm thấy nhân viên mã: " + maNV,
                            "Thông báo",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        continue;
                    }

                    int nhanVienID =
                        Convert.ToInt32(result);

                    // ===== Check nhân viên đã có bảo hiểm =====
                    string sqlCheckNV =
                        "SELECT COUNT(*) FROM BAOHIEM WHERE NhanVien_id=@NhanVien_id";

                    MySqlCommand cmdCheckNV =
                        new MySqlCommand(sqlCheckNV, ConnectData.conn);

                    cmdCheckNV.Parameters.AddWithValue(
                        "@NhanVien_id",
                        nhanVienID);

                    int checkNV =
                        Convert.ToInt32(
                            cmdCheckNV.ExecuteScalar());

                    if (checkNV > 0)
                    {
                        MessageBox.Show(
                            "Nhân viên mã: " + maNV +
                            " đã có bảo hiểm",
                            "Thông báo",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        continue;
                    }

                    // ===== Số bảo hiểm =====
                    string soBaoHiem =
                        Convert.ToString(
                            (range.Cells[i, 7] as Excel.Range).Text);

                    if (string.IsNullOrWhiteSpace(soBaoHiem))
                    {
                        MessageBox.Show(
                            "Dòng " + i +
                            " chưa có số bảo hiểm",
                            "Thông báo",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        continue;
                    }

                    // ===== Check trùng số bảo hiểm =====
                    string sqlCheck =
                        "SELECT COUNT(*) FROM BAOHIEM WHERE So_BaoHiem=@So_BaoHiem";

                    MySqlCommand cmdCheck =
                        new MySqlCommand(sqlCheck, ConnectData.conn);

                    cmdCheck.Parameters.AddWithValue(
                        "@So_BaoHiem",
                        soBaoHiem);

                    int check =
                        Convert.ToInt32(
                            cmdCheck.ExecuteScalar());

                    if (check > 0)
                    {
                        MessageBox.Show(
                            "Số bảo hiểm: " + soBaoHiem +
                            " đã tồn tại",
                            "Thông báo",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        continue;
                    }

                    // ===== Ngày cấp =====
                    DateTime ngayCap = DateTime.Now;

                    object ngayCapValue =
                        (range.Cells[i, 8] as Excel.Range).Value2;

                    if (ngayCapValue != null)
                    {
                        ngayCap =
                            DateTime.FromOADate(
                                Convert.ToDouble(ngayCapValue));
                    }

                    // ===== Nơi cấp =====
                    string noiCap =
                        Convert.ToString(
                            (range.Cells[i, 9] as Excel.Range).Text);

                    // ===== Nơi đăng ký khám =====
                    string noiDK =
                        Convert.ToString(
                            (range.Cells[i, 10] as Excel.Range).Text);

                    // ===== Ghi chú =====
                    string ghiChu =
                        Convert.ToString(
                            (range.Cells[i, 11] as Excel.Range).Text);

                    // ===== Insert =====
                    string sql = @"INSERT INTO BAOHIEM
            (
                NhanVien_id,
                So_BaoHiem,
                NgayCap,
                NoiCap,
                Noi_DK_Kham_Benh,
                GhiChu
            )
            VALUES
            (
                @NhanVien_id,
                @So_BaoHiem,
                @NgayCap,
                @NoiCap,
                @Noi_DK_Kham_Benh,
                @GhiChu
            )";

                    MySqlCommand cmd =
                        new MySqlCommand(sql, ConnectData.conn);

                    cmd.Parameters.AddWithValue(
                        "@NhanVien_id",
                        nhanVienID);

                    cmd.Parameters.AddWithValue(
                        "@So_BaoHiem",
                        soBaoHiem);

                    cmd.Parameters.AddWithValue(
                        "@NgayCap",
                        ngayCap);

                    cmd.Parameters.AddWithValue(
                        "@NoiCap",
                        noiCap);

                    cmd.Parameters.AddWithValue(
                        "@Noi_DK_Kham_Benh",
                        noiDK);

                    cmd.Parameters.AddWithValue(
                        "@GhiChu",
                        ghiChu);

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
            load_baohiem();
            Loadnhanvien();
        }

        private void btn_themBH_Click(object sender, EventArgs e)
        {
            // ===== Kiểm tra =====
            if (lku_nhanvien.EditValue == null)
            {
                MessageBox.Show("Chọn nhân viên");
                lku_nhanvien.Focus();
                return;
            }

            if (txt_sobaohiem.Text.Trim() == "")
            {
                MessageBox.Show("Nhập số bảo hiểm");
                txt_sobaohiem.Focus();
                return;
            }

            if (txt_noicap.Text.Trim() == "")
            {
                MessageBox.Show("Nhập nơi cấp");
                txt_noicap.Focus();
                return;
            }

            if (txt_noikham.Text.Trim() == "")
            {
                MessageBox.Show("Nhập nơi đăng ký khám bệnh");
                txt_noikham.Focus();
                return;
            }

            try
            {
                ConnectData.taoketnoi();

                // ===== Lấy ID nhân viên =====
                int nhanVienID =
                    Convert.ToInt32(
                        lku_nhanvien.EditValue);

                // ===== Check mỗi nhân viên chỉ có 1 bảo hiểm =====
                string sqlCheckNV =
                    "SELECT COUNT(*) FROM BAOHIEM WHERE NhanVien_id=@NhanVien_id";

                MySqlCommand cmdCheckNV =
                    new MySqlCommand(sqlCheckNV, ConnectData.conn);

                cmdCheckNV.Parameters.AddWithValue(
                    "@NhanVien_id",
                    nhanVienID);

                int checkNV =
                    Convert.ToInt32(
                        cmdCheckNV.ExecuteScalar());

                if (checkNV > 0)
                {
                    MessageBox.Show(
                        "Nhân viên này đã có bảo hiểm",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // ===== Check trùng số bảo hiểm =====
                string sqlCheck =
                    "SELECT COUNT(*) FROM BAOHIEM WHERE So_BaoHiem=@So_BaoHiem";

                MySqlCommand cmdCheck =
                    new MySqlCommand(sqlCheck, ConnectData.conn);

                cmdCheck.Parameters.AddWithValue(
                    "@So_BaoHiem",
                    txt_sobaohiem.Text.Trim());

                int check =
                    Convert.ToInt32(
                        cmdCheck.ExecuteScalar());

                if (check > 0)
                {
                    MessageBox.Show(
                        "Số bảo hiểm đã tồn tại");

                    txt_sobaohiem.Focus();
                    return;
                }

                // ===== Insert =====
                string sql = @"INSERT INTO BAOHIEM
        (
            NhanVien_id,
            So_BaoHiem,
            NgayCap,
            NoiCap,
            Noi_DK_Kham_Benh,
            GhiChu
        )
        VALUES
        (
            @NhanVien_id,
            @So_BaoHiem,
            @NgayCap,
            @NoiCap,
            @Noi_DK_Kham_Benh,
            @GhiChu
        )";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@NhanVien_id",
                    nhanVienID);

                cmd.Parameters.AddWithValue(
                    "@So_BaoHiem",
                    txt_sobaohiem.Text.Trim().ToUpper());

                cmd.Parameters.AddWithValue(
                    "@NgayCap",
                    date_cap.Value);

                cmd.Parameters.AddWithValue(
                    "@NoiCap",
                    txt_noicap.Text.Trim());

                cmd.Parameters.AddWithValue(
                    "@Noi_DK_Kham_Benh",
                    txt_noikham.Text.Trim());

                cmd.Parameters.AddWithValue(
                    "@GhiChu",
                    txt_ghichu.Text.Trim());

                cmd.ExecuteNonQuery();

                string hoTen =
                    lku_nhanvien.Text;

                MessageBox.Show(
                    "Thêm bảo hiểm thành công cho:\n"
                    + hoTen,
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);


            }
            catch (MySqlException ex)
            {
                if (ex.Number == 1062)
                {
                    MessageBox.Show(
                        "Dữ liệu đã tồn tại");
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
            load_baohiem();
            Loadnhanvien();
            clearform();
        }

        private void btn_lammoi_Click_1(object sender, EventArgs e)
        {
            
            
            load_baohiem();
            Loadnhanvien();
            clearform();
        }
    }
}