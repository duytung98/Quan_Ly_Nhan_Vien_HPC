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

        // ===== USER LOGIN =====
        public static string CurrentUser = "";

        // ===== PATH LOG =====
        public static string LogFolder =
            Environment.GetFolderPath(
                Environment.SpecialFolder.CommonDocuments)
            + @"\HPC_Logs";

        // ===== GHI LOG =====
        public static void WriteLog(
            string chucNang,
            string hanhDong,
            string noiDung)
        {
            try
            {
                // ===== TẠO THƯ MỤC =====
                if (!Directory.Exists(LogFolder))
                {
                    Directory.CreateDirectory(LogFolder);
                }

                // ===== FILE THEO NGÀY =====
                string fileName =
                    DateTime.Now.ToString("yyyyMMdd")
                    + ".txt";

                string fullPath =
                    Path.Combine(
                        LogFolder,
                        fileName);

                using (StreamWriter sw =
                    new StreamWriter(fullPath, true))
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
