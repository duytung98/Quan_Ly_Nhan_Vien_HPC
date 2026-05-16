using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Quan_Ly_Nhan_Vien_HPC
{
    public class ConnectData
    {
        // 🔹 Khởi tạo kết nối MySQL
        public static MySqlConnection conn = new MySqlConnection();

        // 🔹 Hàm tạo kết nối
        public static void taoketnoi()
        {
            conn.ConnectionString = "Server = localhost;port=3306; Database = quanlynhansu;UId = root;Pwd = root;Pooling = false;Character Set=utf8";

            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Lỗi kết nối MySQL:\n" + ex.Message,
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // 🔹 Đóng kết nối
        public static void dongketnoi()
        {
            if (conn.State == ConnectionState.Open)
                conn.Close();
        }

        // 🔹 Đổ dữ liệu vào DataTable
        public static DataTable getdata(string query)
        {
            DataTable table = new DataTable();

            try
            {
                taoketnoi();

                using (MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn))
                {
                    adapter.Fill(table);
                }

                // 🔹 Thêm cột STT
                if (!table.Columns.Contains("STT"))
                    table.Columns.Add("STT", typeof(int));

                int stt = 1;
                foreach (DataRow row in table.Rows)
                {
                    row["STT"] = stt++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi getdata: " + ex.Message);
            }
            finally
            {
                dongketnoi();
            }

            return table;
        }

        // 🔹 Lấy DataSet
        public static DataSet getdataSet(string query)
        {
            DataSet ds = new DataSet();

            try
            {
                taoketnoi();

                using (MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn))
                {
                    adapter.Fill(ds);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi getdataSet: " + ex.Message);
            }
            finally
            {
                dongketnoi();
            }

            return ds;
        }

        // 🔹 Insert / Update / Delete
        public static void execQuery(string sql)
        {
            try
            {
                taoketnoi();

                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi execQuery: " + ex.Message);
            }
            finally
            {
                dongketnoi();
            }
        }

        // 🔹 ExecuteScalar (COUNT, LOGIN...)
        public static object execScalar(string sql, MySqlParameter[] param = null)
        {
            object result = null;

            try
            {
                taoketnoi();

                MySqlCommand cmd = new MySqlCommand(sql, conn);

                if (param != null)
                    cmd.Parameters.AddRange(param);

                result = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi execScalar: " + ex.Message);
            }
            finally
            {
                dongketnoi();
            }

            return result;
        }
    }
}
