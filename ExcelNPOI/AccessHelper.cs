using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Configuration;

namespace ExcelNPOI
{
    public static partial class AccessHelper
    {
        private static readonly string connectString = ConfigurationManager.ConnectionStrings["accessConnect"].ToString();
        
       /// <summary>
       /// 查询所有数据：不带参数。
       /// </summary>
       /// <param name="strSQL">查询语句</param>
       /// <returns>OleDbDataReader数据类型</returns>
        public static OleDbDataReader SelectAtt(string strSQL, params OleDbParameter[] param)
        {

            IDbConnection dbconnection = new OleDbConnection(connectString);
            OleDbDataReader reader = null;
            using (OleDbCommand cmd = new OleDbCommand(strSQL, dbconnection as OleDbConnection))
            {
                if (dbconnection.State == ConnectionState.Closed)
                {

                        dbconnection.Open();

                }
                if (param != null)
                {
                    cmd.Parameters.AddRange(param);
                }
                reader =  cmd.ExecuteReader(CommandBehavior.CloseConnection); 
            }
            //dbconnection.Close();
            return reader;
        }

        /// <summary>
        /// 增删改
        /// </summary>
        /// <param name="strSQL"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static int Execute(string strSQL, params OleDbParameter[] param)
        {
            using (OleDbConnection dbconnection=new OleDbConnection(connectString))
            {
                if (dbconnection.State == ConnectionState.Closed)
                {
                    dbconnection.Open();
                }
                using (OleDbCommand cmd=new OleDbCommand(strSQL,dbconnection))
                {
                    if (param!=null)
                    {
                        cmd.Parameters.AddRange(param);
                    }

                    
                    return cmd.ExecuteNonQuery();
                    
                }
            }
        }

    }
}
