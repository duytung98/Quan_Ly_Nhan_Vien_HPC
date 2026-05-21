using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public class LogSystem
    {
        // ===== USER ĐĂNG NHẬP =====
        public static string CurrentUser = "";

        // ===== GHI LOG =====
        public static void WriteLog(
            string chucNang,
            string hanhDong,
            string noiDung)
        {
            try
            {
                string folder =
                    Application.StartupPath + @"\Logs";

                // ===== Tạo thư mục Logs =====
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                // ===== File theo ngày =====
                string fileName =
                    DateTime.Now.ToString("yyyyMMdd")
                    + ".txt";

                string path =
                    folder + @"\" + fileName;

                using (StreamWriter sw =
                    new StreamWriter(path, true))
                {
                    sw.WriteLine(
                        "================================================");

                    sw.WriteLine(
                        "Thời gian : "
                        + DateTime.Now.ToString(
                            "dd/MM/yyyy HH:mm:ss"));

                    sw.WriteLine(
                        "Tài khoản : "
                        + CurrentUser);

                    sw.WriteLine(
                        "Chức năng : "
                        + chucNang);

                    sw.WriteLine(
                        "Hành động : "
                        + hanhDong);

                    sw.WriteLine(
                        "Nội dung  : "
                        + noiDung);

                    sw.WriteLine(
                        "================================================");

                    sw.WriteLine();
                }
            }
            catch
            {

            }
        }
    }
}
