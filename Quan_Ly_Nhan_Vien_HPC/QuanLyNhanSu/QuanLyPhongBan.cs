using DevExpress.XtraEditors;
using DevExpress.XtraExport.Helpers;
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
    public partial class QuanLyPhongBan : DevExpress.XtraEditors.XtraForm
    {
        public QuanLyPhongBan()
        {
            InitializeComponent();
        }

        public void loadPhongBan()
        {
            gc_phongban.DataSource = ConnectData.getdata("SELECT * FROM PHONGBAN order by(id) desc;");
        }
        private void QuanLyPhongBan_Load(object sender, EventArgs e)
        {
            loadPhongBan();
            gv_phonban.OptionsBehavior.Editable = true;
            gv_phonban.OptionsView.NewItemRowPosition =DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None;
            GridView gridView = gv_phonban;
            // Định dạng căn giữa chữ trong Header
            gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            // Định dạng căn giữa chữ trong Row
            gridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            btn_sua.Click += Btn_sua_Click;
            btn_xoa.Click += Btn_xoa_Click;
            txt_mapb.KeyDown += txt_mapb_KeyDown;
            txt_tenphong.KeyDown += txt_tenphong_KeyDown;
            txt_ghichu.KeyDown += txt_ghichu_KeyDown;
        }

        private void txt_ghichu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_them.PerformClick();
            }
        }

        private void txt_tenphong_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_ghichu.Focus();
            }
        }

        private void txt_mapb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // chặn tiếng beep + mất chữ

                txt_tenphong.Focus();
            }
        }

        private void Btn_xoa_Click(object sender, EventArgs e)
        {
            if (gv_phonban.FocusedRowHandle < 0)
            {
                MessageBox.Show("Chọn dữ liệu cần xóa");
                return;
            }

            if (MessageBox.Show("Bạn có muốn xóa không?",
                "Thông báo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.No)
                return;

            try
            {
                int id = Convert.ToInt32(
                    gv_phonban.GetFocusedRowCellValue("id"));

                ConnectData.taoketnoi();

                string sql = "DELETE FROM PHONGBAN WHERE id=@id";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue("@id", id);

                cmd.ExecuteNonQuery();

                MessageBox.Show("Xóa thành công");
                LogSystem.WriteLog("Phòng ban", "XÓA", "Xóa phòng ban: " + MaPhongBan + " - " + TenPhongBan + " - " + Login.txt_taikhoan.Text);



            }
            catch (MySqlException ex)
            {
                // Lỗi khóa ngoại
                if (ex.Number == 1451)
                {
                    MessageBox.Show(
                        "Dữ liệu đang được sử dụng, không thể xóa",
                        "Khóa ngoại",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }

                // Trùng khóa chính / unique
                else if (ex.Number == 1062)
                {
                    MessageBox.Show(
                        "Dữ liệu đã tồn tại",
                        "Trùng dữ liệu",
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
            loadPhongBan();
        }

        private void Btn_sua_Click(object sender, EventArgs e)
        {
            if (gv_phonban.FocusedRowHandle < 0)
            {
                MessageBox.Show("Chọn dữ liệu cần sửa");
                return;
            }

            if (MessageBox.Show("Bạn có chắc chắn muốn sửa không?",
                "Thông báo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.No)
                return;

            try
            {
                ConnectData.taoketnoi();

                for (int i = 0; i < gv_phonban.RowCount; i++)
                {
                    object idValue = gv_phonban.GetRowCellValue(i, "id");

                    if (idValue == null)
                        continue;

                    int id = Convert.ToInt32(idValue);

                    string maPhongBan = gv_phonban.GetRowCellValue(i, "MaPhongBan")?.ToString().ToUpper();
                    string tenPhongBan = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(gv_phonban.GetRowCellValue(i, "TenPhongBan")?.ToString().ToLower());
                    string ghiChu = gv_phonban.GetRowCellValue(i, "GhiChu")?.ToString();

                    string sql = @"UPDATE PHONGBAN 
                           SET MaPhongBan=@MaPhongBan,
                               TenPhongBan=@TenPhongBan,
                               GhiChu=@GhiChu
                           WHERE id=@id";

                    MySqlCommand cmd = new MySqlCommand(sql, ConnectData.conn);

                    cmd.Parameters.AddWithValue("@MaPhongBan", maPhongBan);
                    cmd.Parameters.AddWithValue("@TenPhongBan", tenPhongBan);
                    cmd.Parameters.AddWithValue("@GhiChu", ghiChu);
                    cmd.Parameters.AddWithValue("@id", id);

                    cmd.ExecuteNonQuery();
                    LogSystem.WriteLog("Phòng ban", "SỬA", "Sửa phòng ban: " + maPhongBan + " - " + tenPhongBan + " - " + Login.txt_taikhoan.Text);

                }

                MessageBox.Show("Sửa thành công");
            }
            catch (MySqlException ex)
            {
                if (ex.Number == 1062)
                {
                    MessageBox.Show("Mã phòng ban đã tồn tại");
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
            loadPhongBan();
        }

        private void btn_them_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra dữ liệu
                if (txt_mapb.Text.Trim() == "")
                {
                    MessageBox.Show("Vui lòng nhập mã phòng ban");
                    txt_mapb.Focus();
                    return;
                }

                if (txt_tenphong.Text.Trim() == "")
                {
                    MessageBox.Show("Vui lòng nhập tên phòng ban");
                    txt_tenphong.Focus();
                    return;
                }

                ConnectData.taoketnoi();

                // Kiểm tra mã phòng ban tồn tại
                string checkSql =
                    "SELECT COUNT(*) FROM PHONGBAN WHERE MaPhongBan = @mapb";

                MySqlCommand checkCmd =
                    new MySqlCommand(checkSql, ConnectData.conn);

                checkCmd.Parameters.AddWithValue(
                    "@mapb",
                    txt_mapb.Text.Trim()
                );

                int count =
                    Convert.ToInt32(checkCmd.ExecuteScalar());

                if (count > 0)
                {
                    MessageBox.Show(
                        "Mã phòng ban đã tồn tại"
                    );

                    txt_mapb.Focus();
                    return;
                }
                // IN HOA mã phòng ban
                string maPhongBan = txt_mapb.Text.Trim().ToUpper();

                // VIẾT HOA CHỮ CÁI ĐẦU
                string tenPhongBan =
                    System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(txt_tenphong.Text.Trim().ToLower());

                // Insert
                string sql = @"
        INSERT INTO PHONGBAN
        (
            MaPhongBan,
            TenPhongBan,
            GhiChu
        )
        VALUES
        (
            @MaPhongBan,
            @TenPhongBan,
            @GhiChu
        )";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue(
                    "@MaPhongBan",
                    maPhongBan
                );

                cmd.Parameters.AddWithValue(
                    "@TenPhongBan",
                    tenPhongBan
                );

                cmd.Parameters.AddWithValue(
                    "@GhiChu",
                    txt_ghichu.Text.Trim()
                );

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Thêm phòng ban thành công",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                
                LogSystem.WriteLog("Phòng ban", "THÊM", "Thêm phòng ban: " + maPhongBan + " - " + tenPhongBan + " - " + Login.txt_taikhoan.Text);

                // Load lại grid
                

                // Clear dữ liệu
                txt_mapb.Clear();
                txt_tenphong.Clear();
                txt_ghichu.Clear();

                txt_mapb.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Lỗi thêm:\n" + ex.Message
                );
            }
            finally
            {
                ConnectData.dongketnoi();
            }
            loadPhongBan();
        }
        DangNhap Login = (DangNhap)Application.OpenForms["DangNhap"];
        private void btn_lammoi_Click(object sender, EventArgs e)
        {
            txt_mapb.Clear();
            txt_tenphong.Clear();
            txt_ghichu.Clear();
            loadPhongBan();
        }

        private void btn_xuatexcel_Click(object sender, EventArgs e)
        {
            try
            {
                //SaveFileDialog save = new SaveFileDialog();

                //save.Filter = "Excel File|*.xlsx";
                SaveFileDialog save =
                    new SaveFileDialog();

                save.Filter =
                    "Excel Workbook|*.xlsx";

                save.Title =
                    "Xuất Excel";

                save.FileName =
                    "DanhSachThongTinPhongBan.xlsx";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                Excel.Application app = new Excel.Application();

                Excel.Workbook wb = app.Workbooks.Add(Type.Missing);

                Excel.Worksheet ws =
                    (Excel.Worksheet)wb.ActiveSheet;

                ws.Name = "PhongBan";

                int rowCount = gv_phonban.RowCount;

                // Đếm cột hợp lệ
                int totalCol = 0;

                for (int i = 0; i < gv_phonban.Columns.Count; i++)
                {
                    string colName =
                        gv_phonban.Columns[i].FieldName;

                    if (colName == "Sua" || colName == "Xoa")
                        continue;

                    totalCol++;
                }

                // Tiêu đề
                ws.Range[
                    ws.Cells[1, 1],
                    ws.Cells[1, totalCol]
                ].Merge();

                ws.Cells[1, 1] = "DANH SÁCH PHÒNG BAN";

                Excel.Range title =
                    ws.Range[
                        ws.Cells[1, 1],
                        ws.Cells[1, totalCol]
                    ];

                title.Font.Bold = true;
                title.Font.Size = 18;
                title.HorizontalAlignment =
                    Excel.XlHAlign.xlHAlignCenter;

                // Header
                int excelCol = 1;

                for (int i = 0; i < gv_phonban.Columns.Count; i++)
                {
                    string colName =
                        gv_phonban.Columns[i].FieldName;

                    if (colName == "Sua" || colName == "Xoa")
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

                // Data
                for (int i = 0; i < rowCount; i++)
                {
                    excelCol = 1;

                    for (int j = 0; j < gv_phonban.Columns.Count; j++)
                    {
                        string colName =
                            gv_phonban.Columns[j].FieldName;

                        if (colName == "Sua" || colName == "Xoa")
                            continue;

                        ws.Cells[i + 4, excelCol] =
                            gv_phonban.GetRowCellValue(
                                i,
                                gv_phonban.Columns[j]);

                        excelCol++;
                    }
                }

                // Border
                Excel.Range border =
                    ws.Range[
                        ws.Cells[3, 1],
                        ws.Cells[rowCount + 3, totalCol]
                    ];

                border.Borders.LineStyle =
                    Excel.XlLineStyle.xlContinuous;

                border.Borders.Weight =
                    Excel.XlBorderWeight.xlThin;

                // AutoFit
                ws.Columns.AutoFit();

                wb.SaveAs(save.FileName);

                wb.Close();
                app.Quit();

                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);

                MessageBox.Show("Xuất Excel thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_import_Click(object sender, EventArgs e)
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
                        // ===== Mã phòng ban =====
                        string maPhongBan =
                            Convert.ToString(
                                (range.Cells[i, 2] as Excel.Range).Text);

                        if (string.IsNullOrWhiteSpace(maPhongBan))
                            continue;

                        // ===== Tên phòng ban =====
                        string tenPhongBan =
                            Convert.ToString(
                                (range.Cells[i, 3] as Excel.Range).Text);

                        if (string.IsNullOrWhiteSpace(tenPhongBan))
                        {
                            MessageBox.Show(
                                "Dòng "
                                + i
                                + " chưa có tên phòng ban",
                                "Thông báo",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);

                            continue;
                        }

                        // ===== Ghi chú =====
                        string ghiChu =
                            Convert.ToString(
                                (range.Cells[i, 4] as Excel.Range).Text);

                        // ===== Check trùng mã phòng ban =====
                        string sqlCheck =
                            "SELECT COUNT(*) FROM PHONGBAN WHERE MaPhongBan=@MaPhongBan";

                        MySqlCommand cmdCheck =
                            new MySqlCommand(
                                sqlCheck,
                                ConnectData.conn);

                        cmdCheck.Parameters.AddWithValue(
                            "@MaPhongBan",
                            maPhongBan);

                        int check =
                            Convert.ToInt32(
                                cmdCheck.ExecuteScalar());

                        if (check > 0)
                        {
                            MessageBox.Show(
                                "Mã phòng ban "
                                + maPhongBan
                                + " đã tồn tại",
                                "Thông báo",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);

                            continue;
                        }

                        // ===== Insert =====
                       string sql = @"
                        INSERT INTO PHONGBAN
                        (
                            MaPhongBan,
                            TenPhongBan,
                            GhiChu
                        )
                        VALUES
                        (
                            @MaPhongBan,
                            @TenPhongBan,
                            @GhiChu
                        )";

                        MySqlCommand cmd =
                            new MySqlCommand(
                                sql,
                                ConnectData.conn);

                        cmd.Parameters.AddWithValue(
                            "@MaPhongBan",
                            maPhongBan);

                        cmd.Parameters.AddWithValue(
                            "@TenPhongBan",
                            tenPhongBan);

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
                            + ":\n"
                            + exRow.Message);
                    }
                }

                wb.Close(false);
                app.Quit();

                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);

                MessageBox.Show(
                    "Import phòng ban thành công",
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

            loadPhongBan();
        }
    }
}