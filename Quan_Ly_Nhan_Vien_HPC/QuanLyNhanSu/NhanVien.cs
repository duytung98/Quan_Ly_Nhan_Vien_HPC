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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class NhanVien : DevExpress.XtraEditors.XtraForm
    {
        public NhanVien()
        {
            InitializeComponent();
        }

        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            F_ThongTinGiangVien f_ThongTin = new F_ThongTinGiangVien();
            f_ThongTin.ShowDialog();
        }
        public void loadData()
        {
            gc_nhanvien.DataSource = ConnectData.getdata("SELECT * FROM NHANVIEN WHERE DELETEO_BY IS NULL or TrangThai = \"Đang làm việc\" ORDER BY id desc;");
        }
        private void NhanVien_Load(object sender, EventArgs e)
        {

            loadData();
            gv_nhanvien.OptionsBehavior.Editable = true;
            gv_nhanvien.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None;
            GridView gridView = gv_nhanvien;
            // Định dạng căn giữa chữ trong Header
            gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            // Định dạng căn giữa chữ trong Row
            gridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            btnSua1.Click += Btn_sua_Click;
            btnXoa1.Click += Btn_xoa_Click;
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

        private void Btn_xoa_Click(object sender, EventArgs e)
        {
            if (gv_nhanvien.FocusedRowHandle < 0)
            {
                MessageBox.Show("Chọn nhân viên cần xóa");
                return;
            }

            string maNV =gv_nhanvien.GetFocusedRowCellValue("MaNV").ToString();

            string hoTen =gv_nhanvien.GetFocusedRowCellValue("HoDem")+ " " +gv_nhanvien.GetFocusedRowCellValue("Ten");

            if (MessageBox.Show(
                "Bạn có chắc muốn xóa nhân viên:\n"
                + maNV + " - " + hoTen + " ?",
                "Thông báo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.No)
                return;

            try
            {
                ConnectData.taoketnoi();

                int id = Convert.ToInt32(
                    gv_nhanvien.GetFocusedRowCellValue("id"));

                string sql = @"UPDATE NHANVIEN SET
                       DELETEO_BY=@DELETEO_BY,
                       DELETEO_DATE=@DELETEO_DATE,
                          TrangThai='Đã nghỉ việc'
                       WHERE id=@id";

                MySqlCommand cmd =
                    new MySqlCommand(sql, ConnectData.conn);

                cmd.Parameters.AddWithValue("@DELETEO_BY",Login.txt_taikhoan.Text);

                cmd.Parameters.AddWithValue("@DELETEO_DATE",DateTime.Now);

                cmd.Parameters.AddWithValue("@id",id);

                cmd.ExecuteNonQuery();

                MessageBox.Show(
                    "Xóa thành công nhân viên:\n"
                    + maNV + " - " + hoTen,
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
        DangNhap Login = (DangNhap)Application.OpenForms["DangNhap"];
        private void Btn_sua_Click(object sender, EventArgs e)
        {
            if (gv_nhanvien.FocusedRowHandle < 0)
            {
                MessageBox.Show("Chọn nhân viên cần sửa");
                return;
            }

            string maNV = gv_nhanvien.GetFocusedRowCellValue("MaNV").ToString();

            string hoTen = gv_nhanvien.GetFocusedRowCellValue("HoDem") + " " + gv_nhanvien.GetFocusedRowCellValue("Ten");
            if (MessageBox.Show("Bạn có chắc muốn sửa thông tin nhân viên: \n" + maNV + " - " + hoTen,
                "Thông báo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.No)
                return;

            try
            {
                ConnectData.taoketnoi();

                for (int i = 0; i < gv_nhanvien.RowCount; i++)
                {
                    object idValue =
                        gv_nhanvien.GetRowCellValue(i, "id");

                    if (idValue == null)
                        continue;

                    int id = Convert.ToInt32(idValue);

                    string hoDem =
                        System.Globalization.CultureInfo
                        .CurrentCulture.TextInfo
                        .ToTitleCase(
                            gv_nhanvien.GetRowCellValue(i, "HoDem")
                            ?.ToString()
                            .ToLower());

                    string ten =
                        System.Globalization.CultureInfo
                        .CurrentCulture.TextInfo
                        .ToTitleCase(
                            gv_nhanvien.GetRowCellValue(i, "Ten")
                            ?.ToString()
                            .ToLower());

                    string cccd =
                        gv_nhanvien.GetRowCellValue(i, "CCCD")
                        ?.ToString();

                    string email =
                        gv_nhanvien.GetRowCellValue(i, "Email")
                        ?.ToString();

                    string sdt =
                        gv_nhanvien.GetRowCellValue(i, "SDT")
                        ?.ToString();

                    // ===== Check CCCD =====
                    string sqlCCCD =
                        "SELECT COUNT(*) FROM NHANVIEN WHERE CCCD=@CCCD AND id<>@id";

                    MySqlCommand cmdCCCD =
                        new MySqlCommand(sqlCCCD, ConnectData.conn);

                    cmdCCCD.Parameters.AddWithValue("@CCCD", cccd);
                    cmdCCCD.Parameters.AddWithValue("@id", id);

                    int checkCCCD =
                        Convert.ToInt32(cmdCCCD.ExecuteScalar());

                    if (checkCCCD > 0)
                    {
                        MessageBox.Show("CCCD đã tồn tại");
                        return;
                    }

                    // ===== Check Email =====
                    string sqlEmail =
                        "SELECT COUNT(*) FROM NHANVIEN WHERE Email=@Email AND id<>@id";

                    MySqlCommand cmdEmail =
                        new MySqlCommand(sqlEmail, ConnectData.conn);

                    cmdEmail.Parameters.AddWithValue("@Email", email);
                    cmdEmail.Parameters.AddWithValue("@id", id);

                    int checkEmail =
                        Convert.ToInt32(cmdEmail.ExecuteScalar());

                    if (checkEmail > 0)
                    {
                        MessageBox.Show("Email đã tồn tại");
                        return;
                    }

                    // ===== Check SDT =====
                    

                    // ===== Update =====
                    string sql = @"UPDATE NHANVIEN SET
            HoDem=@HoDem,
            Ten=@Ten,
            Password=@Password,
            NgaySinh=@NgaySinh,
            GioiTinh=@GioiTinh,
            DanToc=@DanToc,
            QuocTich=@QuocTich,
            TonGiao=@TonGiao,
            QueQuan=@QueQuan,
            DiaChiThuongChu=@DiaChiThuongChu,
            CCCD=@CCCD,
            NoiCapCCCD=@NoiCapCCCD,
            NgayCapCCCD=@NgayCapCCCD,
            SDT=@SDT,
            Email=@Email,
            TrinhDoGiaoDucPhoThong=@TrinhDoGiaoDucPhoThong,
            TrinhDoChuyenMon=@TrinhDoChuyenMon,
            HocHamHocVi=@HocHamHocVi,
            ChungChiKhac=@ChungChiKhac,
            HinhAnh=@HinhAnh,
            MaSoThue=@MaSoThue,
            TrangThai=@TrangThai,
            GhiChu=@GhiChu,
            UPDATEO_BY=@UPDATEO_BY,
            UPDATEO_DATE=@UPDATEO_DATE
            WHERE id=@id";

                    MySqlCommand cmd =
                        new MySqlCommand(sql, ConnectData.conn);

                    cmd.Parameters.AddWithValue("@HoDem", hoDem);
                    cmd.Parameters.AddWithValue("@Ten", ten);

                    cmd.Parameters.AddWithValue("@Password",gv_nhanvien.GetRowCellValue(i, "Password"));

                    cmd.Parameters.AddWithValue("@NgaySinh",gv_nhanvien.GetRowCellValue(i, "NgaySinh"));

                    cmd.Parameters.AddWithValue("@GioiTinh",gv_nhanvien.GetRowCellValue(i, "GioiTinh"));

                    cmd.Parameters.AddWithValue("@DanToc",gv_nhanvien.GetRowCellValue(i, "DanToc"));

                    cmd.Parameters.AddWithValue("@QuocTich",gv_nhanvien.GetRowCellValue(i, "QuocTich"));

                    cmd.Parameters.AddWithValue("@TonGiao",gv_nhanvien.GetRowCellValue(i, "TonGiao"));

                    cmd.Parameters.AddWithValue("@QueQuan",gv_nhanvien.GetRowCellValue(i, "QueQuan"));

                    cmd.Parameters.AddWithValue("@DiaChiThuongChu",gv_nhanvien.GetRowCellValue(i, "DiaChiThuongChu"));

                    cmd.Parameters.AddWithValue("@CCCD", cccd);

                    cmd.Parameters.AddWithValue("@NoiCapCCCD",gv_nhanvien.GetRowCellValue(i, "NoiCapCCCD"));

                    cmd.Parameters.AddWithValue("@NgayCapCCCD",gv_nhanvien.GetRowCellValue(i, "NgayCapCCCD"));

                    cmd.Parameters.AddWithValue("@SDT", gv_nhanvien.GetRowCellValue(i, "SDT"));

                    cmd.Parameters.AddWithValue("@Email", email);

                    cmd.Parameters.AddWithValue("@TrinhDoGiaoDucPhoThong",gv_nhanvien.GetRowCellValue(i, "TrinhDoGiaoDucPhoThong"));

                    cmd.Parameters.AddWithValue("@TrinhDoChuyenMon",gv_nhanvien.GetRowCellValue(i, "TrinhDoChuyenMon"));

                    cmd.Parameters.AddWithValue("@HocHamHocVi",gv_nhanvien.GetRowCellValue(i, "HocHamHocVi"));

                    cmd.Parameters.AddWithValue("@ChungChiKhac", gv_nhanvien.GetRowCellValue(i, "ChungChiKhac"));

                    cmd.Parameters.AddWithValue("@HinhAnh", gv_nhanvien.GetRowCellValue(i, "HinhAnh"));

                    cmd.Parameters.AddWithValue("@MaSoThue", gv_nhanvien.GetRowCellValue(i, "MaSoThue"));

                    cmd.Parameters.AddWithValue("@TrangThai", gv_nhanvien.GetRowCellValue(i, "TrangThai"));

                    cmd.Parameters.AddWithValue("@GhiChu",gv_nhanvien.GetRowCellValue(i, "GhiChu"));

                    cmd.Parameters.AddWithValue("@UPDATEO_BY", Login.txt_taikhoan.Text);

                    cmd.Parameters.AddWithValue("@UPDATEO_DATE",DateTime.Now);

                    cmd.Parameters.AddWithValue("@id", id);

                    cmd.ExecuteNonQuery();
                }

                //string maNV =gv_nhanvien.GetFocusedRowCellValue("MaNV").ToString();

                //string hoTen =gv_nhanvien.GetFocusedRowCellValue("HoDem")+ " " +gv_nhanvien.GetFocusedRowCellValue("Ten");

                MessageBox.Show("Sửa thành công thông tin nhân viên:\n"+ maNV + " - " + hoTen,"Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);


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
            loadData();
        }

        private void btn_import_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loadData();
        }

        private void btn_xuat_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
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
                    "DanhSachThongTinNhanVien.xlsx";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                Excel.Application app =
                    new Excel.Application();

                Excel.Workbook wb =
                    app.Workbooks.Add(Type.Missing);

                Excel.Worksheet ws =
                    (Excel.Worksheet)wb.ActiveSheet;

                ws.Name = "NhanVien";

                int rowCount = gv_nhanvien.RowCount;

                int totalCol = 0;

                // ===== Đếm cột =====
                for (int i = 0; i < gv_nhanvien.Columns.Count; i++)
                {
                    string colName =
                        gv_nhanvien.Columns[i].FieldName;

                    if (colName == "Sua"
                        || colName == "Xoa"
                        || colName == "Password"
                        || colName == "DELETEO_BY"
                        || colName == "DELETEO_DATE")
                        continue;

                    totalCol++;
                }

                // ===== Tiêu đề =====
                ws.Range[
                    ws.Cells[1, 1],
                    ws.Cells[1, totalCol]
                ].Merge();

                ws.Cells[1, 1] =
                    "DANH SÁCH NHÂN VIÊN";

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

                for (int i = 0; i < gv_nhanvien.Columns.Count; i++)
                {
                    string colName =
                        gv_nhanvien.Columns[i].FieldName;

                    if (colName == "Sua"
                        || colName == "Xoa"
                        || colName == "Password"
                        || colName == "DELETEO_BY"
                        || colName == "DELETEO_DATE")
                        continue;

                    ws.Cells[3, excelCol] =
                        gv_nhanvien.Columns[i].Caption;

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

                    for (int j = 0; j < gv_nhanvien.Columns.Count; j++)
                    {
                        string colName =
                            gv_nhanvien.Columns[j].FieldName;

                        if (colName == "Sua"
                            || colName == "Xoa"
                            || colName == "Password"
                            || colName == "DELETEO_BY"
                            || colName == "DELETEO_DATE")
                            continue;

                        object value =
                            gv_nhanvien.GetRowCellValue(
                                i,
                                gv_nhanvien.Columns[j]);

                        // ===== Ép kiểu text =====
                        if (colName == "CCCD"
                            || colName == "MaSoThue"
                            || colName == "SDT"
                            || colName == "TrinhDoGiaoDucPhoThong")
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

        private void btn_import_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
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

                // Bắt đầu từ dòng 4
                for (int i = 4; i <= rowCount; i++)
                {
                    string maNV =
                        Convert.ToString(
                            (range.Cells[i, 2] as Excel.Range).Text);

                    if (string.IsNullOrWhiteSpace(maNV))
                        continue;

                    // ===== Check trùng mã =====
                    string sqlCheck =
                        "SELECT COUNT(*) FROM NHANVIEN WHERE MaNV=@MaNV";

                    MySqlCommand cmdCheck =
                        new MySqlCommand(sqlCheck, ConnectData.conn);

                    cmdCheck.Parameters.AddWithValue(
                        "@MaNV",
                        maNV);

                    int check =
                        Convert.ToInt32(cmdCheck.ExecuteScalar());

                    if (check > 0)
                        continue;

                    string hoDem =
                        Convert.ToString(
                            (range.Cells[i, 3] as Excel.Range).Text);

                    string ten =
                        Convert.ToString(
                            (range.Cells[i, 4] as Excel.Range).Text);

                    DateTime ngaySinh =
                        Convert.ToDateTime(
                            (range.Cells[i, 5] as Excel.Range).Value);

                    string gioiTinh =
                        Convert.ToString(
                            (range.Cells[i, 6] as Excel.Range).Text);

                    string danToc =
                        Convert.ToString(
                            (range.Cells[i, 7] as Excel.Range).Text);

                    string quocTich =
                        Convert.ToString(
                            (range.Cells[i, 8] as Excel.Range).Text);

                    string tonGiao =
                        Convert.ToString(
                            (range.Cells[i, 9] as Excel.Range).Text);

                    string diaChi =
                        Convert.ToString(
                            (range.Cells[i, 10] as Excel.Range).Text);

                    string queQuan =
                        Convert.ToString(
                            (range.Cells[i, 11] as Excel.Range).Text);

                    string cccd =
                        Convert.ToString(
                            (range.Cells[i, 12] as Excel.Range).Text);

                    string noiCap =
                        Convert.ToString(
                            (range.Cells[i, 13] as Excel.Range).Text);

                    DateTime ngayCap =
                        Convert.ToDateTime(
                            (range.Cells[i, 14] as Excel.Range).Value);

                    string sdt =
                        Convert.ToString(
                            (range.Cells[i, 15] as Excel.Range).Text);

                    string email =
                        Convert.ToString(
                            (range.Cells[i, 16] as Excel.Range).Text);

                    string chuyenMon =
                        Convert.ToString(
                            (range.Cells[i, 17] as Excel.Range).Text);

                    string trinhDo =
                        Convert.ToString(
                            (range.Cells[i, 18] as Excel.Range).Text);

                    string hocHam =
                        Convert.ToString(
                            (range.Cells[i, 19] as Excel.Range).Text);

                    string chungChi =
                        Convert.ToString(
                            (range.Cells[i, 20] as Excel.Range).Text);

                    string maSoThue =
                        Convert.ToString(
                            (range.Cells[i, 21] as Excel.Range).Text);

                    string ghiChu =
                        Convert.ToString(
                            (range.Cells[i, 22] as Excel.Range).Text);
                    string pass = "12345678";
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
                @MaSoThue,
                @TrangThai,
                @GhiChu,
                @CREATEO_BY,
                @CREATEO_DATE
            )";

                    MySqlCommand cmd =
                        new MySqlCommand(sql, ConnectData.conn);

                    cmd.Parameters.AddWithValue("@MaNV", maNV);
                    cmd.Parameters.AddWithValue("@HoDem", hoDem);
                    cmd.Parameters.AddWithValue("@Ten", ten);

                    // Password mặc định
                    cmd.Parameters.AddWithValue(
                        "@Password",
                         BCrypt.Net.BCrypt.EnhancedHashPassword(pass, 10));

                    cmd.Parameters.AddWithValue("@NgaySinh", ngaySinh);
                    cmd.Parameters.AddWithValue("@GioiTinh", gioiTinh);
                    cmd.Parameters.AddWithValue("@DanToc", danToc);
                    cmd.Parameters.AddWithValue("@QuocTich", quocTich);
                    cmd.Parameters.AddWithValue("@TonGiao", tonGiao);
                    cmd.Parameters.AddWithValue("@QueQuan", queQuan);
                    cmd.Parameters.AddWithValue("@DiaChiThuongChu", diaChi);
                    cmd.Parameters.AddWithValue("@CCCD", cccd);
                    cmd.Parameters.AddWithValue("@NoiCapCCCD", noiCap);
                    cmd.Parameters.AddWithValue("@NgayCapCCCD", ngayCap);
                    cmd.Parameters.AddWithValue("@SDT", sdt);
                    cmd.Parameters.AddWithValue("@Email", email);

                    cmd.Parameters.AddWithValue(
                        "@TrinhDoGiaoDucPhoThong",
                        trinhDo);

                    cmd.Parameters.AddWithValue(
                        "@TrinhDoChuyenMon",
                        chuyenMon);

                    cmd.Parameters.AddWithValue(
                        "@HocHamHocVi",
                        hocHam);

                    cmd.Parameters.AddWithValue(
                        "@ChungChiKhac",
                        chungChi);

                    cmd.Parameters.AddWithValue(
                        "@MaSoThue",
                        maSoThue);

                    cmd.Parameters.AddWithValue(
                        "@TrangThai",
                        "Đang làm việc");

                    cmd.Parameters.AddWithValue(
                        "@GhiChu",
                        ghiChu);

                    cmd.Parameters.AddWithValue(
                        "@CREATEO_BY",
                        Login.txt_taikhoan.Text);

                    cmd.Parameters.AddWithValue(
                        "@CREATEO_DATE",
                        DateTime.Now);

                    cmd.ExecuteNonQuery();
                }

                wb.Close(false);
                app.Quit();

                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);

                MessageBox.Show("Import Excel thành công");

                
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