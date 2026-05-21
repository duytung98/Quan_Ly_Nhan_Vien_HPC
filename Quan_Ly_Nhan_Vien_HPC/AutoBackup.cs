using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public class AutoBackup
    {
        public static void BackupDatabase()
        {
            try
            {
                // ===== Chỉ backup ngày 1 =====
                if (DateTime.Now.Day != 1)
                    return;

                // ===== FOLDER BACKUP =====
                string folder =
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.CommonDocuments)
                    + @"\HPC_Backup";

                // ===== TẠO THƯ MỤC =====
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                // ===== TÊN FILE =====
                string fileName =
                    "Backup_"
                    + DateTime.Now.ToString("yyyyMMdd")
                    + ".sql";

                string fullPath =
                    Path.Combine(folder, fileName);

                // ===== NẾU ĐÃ BACKUP =====
                if (File.Exists(fullPath))
                    return;

                // ===== MYSQL INFO =====
                string host =
                    ConnectData.server;

                string port =
                    ConnectData.port;

                string database =
                    ConnectData.database;

                string user =
                    ConnectData.user;

                string password =
                    ConnectData.password;

                // ===== MYSQLDUMP =====
                string mysqldump =
                    @"C:\xampp\mysql\bin\mysqldump.exe";

                // ===== CHECK MYSQLDUMP =====
                if (!File.Exists(mysqldump))
                {
                    MessageBox.Show(
                        "Không tìm thấy mysqldump.exe");

                    return;
                }

                // ===== COMMAND =====
                string arguments =
                    $"-h {host} -P {port} -u {user} -p{password} {database} --result-file=\"{fullPath}\"";

                ProcessStartInfo psi =
                    new ProcessStartInfo();

                psi.FileName =
                    mysqldump;

                psi.Arguments =
                    arguments;

                psi.WindowStyle =
                    ProcessWindowStyle.Hidden;

                psi.CreateNoWindow =
                    true;

                psi.UseShellExecute =
                    false;

                Process process =
                    new Process();

                process.StartInfo =
                    psi;

                process.Start();

                process.WaitForExit();

                // ===== GHI LOG =====
                LogSystem.WriteLog(
                    "Backup",
                    "AUTO BACKUP",
                    "Backup tự động ngày 1");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
