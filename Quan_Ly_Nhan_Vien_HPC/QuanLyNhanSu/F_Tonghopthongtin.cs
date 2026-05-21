using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public partial class F_Tonghopthongtin : DevExpress.XtraEditors.XtraForm
    {
        public F_Tonghopthongtin()
        {
            InitializeComponent();
        }

        private void F_Tonghopthongtin_Load(object sender, EventArgs e)
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

        private void loadData()
        {
            gc_nhanvien.DataSource = ConnectData.getdata("SELECT \r\n    nv.*,\r\n    -- Phòng ban\r\n    pb.MaPhongBan,\r\n    pb.TenPhongBan,\r\n    -- Hợp đồng\r\n    hd.So_HopDong,\r\n    hd.NgayBD,\r\n    hd.LoaiHopDong,\r\n    hd.ChucVu,\r\n    hd.LuongCB,\r\n    hd.HeSoLuong,\r\n    -- Bảo hiểm\r\n    bh.So_BaoHiem,\r\n    bh.NgayCap ,\r\n    bh.NoiCap,\r\n    bh.Noi_DK_Kham_Benh,\r\n    -- Thông tin gia đình\r\n    gd.HoTenBo,\r\n    gd.NgaySinhBo,\r\n    gd.NgheNghiepBo,\r\n    gd.SDTBo,\r\n    gd.HoTenMe,\r\n    gd.NgaySinhMe,\r\n    gd.NgheNghiepMe,\r\n    gd.SDTMe,\r\n    gd.HoTenVo_Chong,\r\n    gd.SDTVo_Chong,\r\n    gd.NgaySinhVo_Chong,\r\n    gd.NgheNghiepVo_Chong,\r\n    gd.HoTenCon1,\r\n    gd.NgaySinhCon1,\r\n    gd.HocVanCon1,\r\n    gd.HoTenCon2,\r\n    gd.NgaySinhCon2,\r\n    gd.HocVanCon2,\r\n    gd.HoTenCon3,\r\n    gd.NgaySinhCon3,\r\n    gd.HocVanCon3,\r\n    -- Thông tin Đảng\r\n    td.NGAYVAODOAN,\r\n    td.NGAYVAODANG,\r\n    td.NgayChinhThucVaoDang\r\nFROM NHANVIEN nv\r\nLEFT JOIN HOPDONG hd\r\n    ON nv.id = hd.NhanVien_id\r\nLEFT JOIN PHONGBAN pb\r\n    ON hd.PhongBan_ID = pb.id\r\nLEFT JOIN BAOHIEM bh\r\n    ON nv.id = bh.NhanVien_id\r\nLEFT JOIN THONGTINGIADINH gd\r\n    ON nv.id = gd.NhanVien_id\r\nLEFT JOIN THONGTINDANG td\r\n    ON nv.id = td.NhanVien_id\r\nWHERE nv.TrangThai = 'Đang làm việc'\r\nAND nv.DELETEO_BY IS NULL;");

        }

        private void Btn_xoa_Click(object sender, EventArgs e)
        {
            
        }

        private void Btn_sua_Click(object sender, EventArgs e)
        {
            
        }

        private void btn_xuat_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (gv_nhanvien.RowCount <= 0)
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
                    "DanhSachNhanVien.xlsx";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                // ===== KHỞI TẠO EXCEL =====
                Microsoft.Office.Interop.Excel.Application app =
                    new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook wb =
                    app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet ws =
                    (Microsoft.Office.Interop.Excel.Worksheet)
                    wb.ActiveSheet;

                ws.Name = "NHANVIEN";

                int cot =
                    gv_nhanvien.Columns.Count;

                int dong =
                    gv_nhanvien.RowCount;

                // ===== TIÊU ĐỀ =====
                Microsoft.Office.Interop.Excel.Range title =
                    ws.Range["A1", "AZ1"];

                title.Merge();

                title.Value =
                    "DANH SÁCH THÔNG TIN NHÂN VIÊN TRƯỜNG CAO ĐẲNG CÔNG NGHỆ BÁCH KHOA HÀ NỘI";

                title.Font.Bold = true;

                title.Font.Size = 18;

                title.HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // ===== HEADER =====
                int colExcel = 1;

                for (int i = 0; i < cot; i++)
                {
                    string tenCot =
                        gv_nhanvien.Columns[i].FieldName;

                    // ===== BỎ CỘT ID =====
                    if (tenCot == "id")
                        continue;

                    ws.Cells[3, colExcel] =
                        gv_nhanvien.Columns[i].Caption;

                    // ===== STYLE =====
                    ws.Cells[3, colExcel].Font.Bold =
                        true;

                    ws.Cells[3, colExcel].Borders.LineStyle =
                        1;

                    ws.Cells[3, colExcel].Interior.Color =
                        System.Drawing.Color.LightGray;

                    // ===== FORMAT TEXT =====
                    ws.Columns[colExcel].NumberFormat =
                        "@";

                    colExcel++;
                }

                // ===== DỮ LIỆU =====
                for (int i = 0; i < dong; i++)
                {
                    colExcel = 1;

                    for (int j = 0; j < cot; j++)
                    {
                        string tenCot =
                            gv_nhanvien.Columns[j].FieldName;

                        // ===== BỎ CỘT ID =====
                        if (tenCot == "id")
                            continue;

                        object value =
                            gv_nhanvien.GetRowCellValue(
                                i,
                                gv_nhanvien.Columns[j]);

                        // ===== NULL =====
                        if (value == null ||
                            value == DBNull.Value)
                        {
                            ws.Cells[i + 4, colExcel] = "";
                        }
                        else
                        {
                            // ===== CHECK CỘT NGÀY =====
                            if
                            (
                                tenCot == "NgaySinh"
                                || tenCot == "NgayCapCCCD"
                                || tenCot == "CREATEO_DATE"
                                || tenCot == "UPDATEO_DATE"
                                || tenCot == "DELETEO_DATE"
                                || tenCot == "NgayBD"
                                || tenCot == "NgayChinhThucVaoDang"
                                || tenCot == "NGAYVAODANG"
                                || tenCot == "NGAYVAODOAN"
                                || tenCot == "NgaySinhBo"
                                || tenCot == "NgaySinhMe"
                                || tenCot == "NgaySinhVo_Chong"
                                || tenCot == "NgaySinhCon1"
                                || tenCot == "NgaySinhCon2"
                                || tenCot == "NgaySinhCon3"
                                || tenCot == "NgayCap"
                            )
                            {
                                ws.Cells[i + 4, colExcel] =
                                    Convert.ToDateTime(value)
                                    .ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                // ===== GIỮ NGUYÊN DỮ LIỆU =====
                                ws.Cells[i + 4, colExcel] =
                                    value.ToString();
                            }
                        }

                        // ===== BORDER =====
                        ws.Cells[i + 4, colExcel]
                            .Borders.LineStyle = 1;

                        colExcel++;
                    }
                }

                // ===== CĂN GIỮA =====
                ws.Range["A3",
                    ws.Cells[dong + 3, colExcel - 1]]
                    .HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // ===== AUTOFIT =====
                ws.Columns.AutoFit();

                // ===== LƯU FILE =====
                wb.SaveAs(save.FileName);

                wb.Close(false);

                app.Quit();

                // ===== GIẢI PHÓNG BỘ NHỚ =====
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                MessageBox.Show(
                    "Xuất Excel thành công",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // ===== MỞ FILE =====
                System.Diagnostics.Process.Start(save.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private void btn_lammoi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loadData();
        }
    }
}