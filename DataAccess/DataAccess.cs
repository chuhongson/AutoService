using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace DataAccess
{
    public class DataAccess
    {
        //private string _connectString = @"Data Source=server;Initial Catalog=CFNGON_FBOHRMSP203_APP;User ID=tvud;Password=123456789";
        private string _connectString = "Data Source=Server;Initial Catalog=ANDIEN_FBOSP191_App;Application Name=%UserID;Uid=tvud;Pwd=123456789;";
        //private string _connectString = "Data Source=UTTHAO;Initial Catalog=Release_FBOHRMSP221_App;Application Name=%UserID;Uid=sa;Pwd=123456;";
        private SqlConnection conn;
        private SqlCommand cmd;


        /// <summary>
        /// Mo ket noi
        /// </summary>
        public void Open()
        {
            if (conn == null)
            {
                conn = new SqlConnection();
                conn.ConnectionString = _connectString;
            }

            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
        }

        public void Close()
        {
            if (conn != null && conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }

        /// <summary>
        /// sử dụng khi dùng câu truy vấn lấy 1 giá trị duy nhất là đối tượng
        /// </summary>
        /// <param name="sql"></param>
        /// <returns>object</returns>
        public object GetObjectValue(string sql)
        {
            Open();
            cmd = new SqlCommand(sql, conn);
            // Trả về giá trị đầu tiên trả về của câu truy vấn
            object result = cmd.ExecuteScalar();
            Close();

            return result;
        }

        /// <summary>
        /// sử dụng khi dùng câu truy vấn lấy 1 giá trị duy nhất là đối tượng
        /// </summary>
        /// <param name="sql"></param>
        /// <returns>object</returns>
        public object GetObjectValue(string sql, SqlParameter[] param)
        {
            Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddRange(param);
            // Trả về giá trị đầu tiên trả về của câu truy vấn
            object result = cmd.ExecuteScalar();
            Close();

            return result;
        }

        /// <summary>
        /// Lấy dữ liệu từ câu truy vấn
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataTable GetData(string sql)
        {
            Open();
            cmd = new SqlCommand(sql, conn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            Close();
            return dt;
        }

        /// <summary>
        /// Lấy dữ liệu từ câu truy vấn
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataTable GetData(string sql, SqlParameter[] param)
        {
            Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddRange(param);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            Close();
            return dt;
        }

        /// <summary>
        /// Trả về số hàng đã bị ảnh hưởng (thêm/bớt/thay đổi)
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public int Execute(string sql)
        {
            Open();
            cmd = new SqlCommand(sql, conn);
            int countEffect = cmd.ExecuteNonQuery();
            Close();

            return countEffect;
        }

        /// <summary>
        /// Trả về số hàng đã bị ảnh hưởng (thêm/bớt/thay đổi)
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public int Execute(string sql, SqlParameter[] param)
        {
            Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddRange(param);
            int countEffect = cmd.ExecuteNonQuery();
            Close();

            return countEffect;
        }

        /// <summary>
        /// when call proc then return DtaTable
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataTable ExecuteProc(string ProcName)
        {
            Open();
            cmd = new SqlCommand(ProcName, conn);
            cmd.CommandType = CommandType.StoredProcedure;
            DataTable dt = new DataTable();
            SqlDataAdapter sqlData = new SqlDataAdapter(cmd);
            sqlData.Fill(dt);
            Close();
            return dt;
        }

        /// <summary>
        /// when call proc with pramater then return DtaTable
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataTable ExecuteProc(string ProcName, DateTime dFrom, DateTime dTo)
        {
            Open();
            DataTable dt = new DataTable();
            cmd = new SqlCommand(ProcName, conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@dFrom", SqlDbType.DateTime).Value = dFrom;
            cmd.Parameters.AddWithValue("@dTo", SqlDbType.DateTime).Value = dTo;
            
            SqlDataAdapter sqlData = new SqlDataAdapter(cmd);
            sqlData.Fill(dt);
            Close();
            return dt;
        }
    }
}
