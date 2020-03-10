using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelNPOI
{
    public static class SQLHelper
    {
        //private static readonly string connectString = ConfigurationManager.ConnectionStrings["SQLConnect"].ToString();
        private static readonly string connectString = "Data Source=192.168.0.158;Initial Catalog=ZKTime;User ID=sa;Password=hrm5w2efa;MultipleActiveResultSets=true";

        /// <summary>
        /// 方法：按条件查询数据
        /// </summary>
        /// <param name="strSQL">sql语句</param>
        /// <param name="param">可变参数</param>
        /// <returns></returns>
        public static SqlDataReader SelectAtt(string strSQL, params SqlParameter[] param)
        {
            IDbConnection dbconnection = new SqlConnection(connectString);
            SqlDataReader reader = null;
            using (SqlCommand cmd=new SqlCommand(strSQL,dbconnection as SqlConnection))
            {
                if (dbconnection.State == ConnectionState.Closed)
                {

                    dbconnection.Open();

                }
                if (param != null)
                {
                    cmd.Parameters.AddRange(param);
                }
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            return reader;
        }

        /// <summary>
        /// 方法：增删改
        /// </summary>
        /// <param name="strSQL"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static int Execute(string strSQL, params SqlParameter[] param)
        {
            using (SqlConnection dbconnection = new SqlConnection(connectString))
            {
                if (dbconnection.State == ConnectionState.Closed)
                {
                    dbconnection.Open();
                }
                using (SqlCommand cmd = new SqlCommand(strSQL, dbconnection))
                {
                    if (param != null)
                    {
                        cmd.Parameters.AddRange(param);
                    }


                    return cmd.ExecuteNonQuery();

                }
            }
        }

        public static object ExecuteScalar(string strSQL,params SqlParameter[] param)
        {
            using (SqlConnection dbconnection = new SqlConnection(connectString))
            {
                if (dbconnection.State == ConnectionState.Closed)
                {
                    dbconnection.Open();
                }
                using (SqlCommand cmd = new SqlCommand(strSQL, dbconnection))
                {
                    if (param != null)
                    {
                        cmd.Parameters.AddRange(param);
                    }


                    return cmd.ExecuteScalar();
                }
            }
        }

        #region 判断数据库表是否存在，通过指定专用的连接字符串，执行一个不需要返回值的SqlCommand命令。
        /// <summary>
        /// 判断数据库表是否存在，返回页头，通过指定专用的连接字符串，执行一个不需要返回值的SqlCommand命令。
        /// </summary>
        /// <param name="tablename">bhtsoft表</param>
        /// <returns></returns>
        public static bool CheckExistsTable(string tablename)
        {
            String tableNameStr = "select count(1) from sysobjects where name = '" + tablename + "'";
            using (SqlConnection con = new SqlConnection(connectString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(tableNameStr, con);
                int result = Convert.ToInt32(cmd.ExecuteScalar());
                if (result == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
        #endregion

    }
}
