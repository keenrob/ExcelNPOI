using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ExcelNPOI
{
    public partial class TblAttListDal
    {

        /// <summary>
        /// 方法：按日期条件生成表tblAttSource
        /// </summary>
        /// <param name="dateB">开始日期</param>
        /// <param name="dateE">结束日期</param>
        /// <returns></returns>
        private int GetExcute(DateTime dateB, DateTime dateE)
        {
            //在Access中，CDate(Format([checktime], 'yyyy-mm-dd')) AS 日期
            string strSQL = "SELECT checkinout.USERID, USERINFO.Name, DEPARTMENTS.DEPTNAME, CONVERT(VARCHAR(10),CHECKTIME,120) AS 日期, checkinout.checktime" +
                            " INTO tblAttSource " +
                            "FROM(USERINFO INNER JOIN checkinout ON USERINFO.USERID = checkinout.USERID) INNER JOIN DEPARTMENTS ON USERINFO.DEFAULTDEPTID = DEPARTMENTS.DEPTID " +
                            "WHERE(checkinout.checktime >= @dateB And checkinout.checktime < @dateE)";
            return SQLHelper.Execute(strSQL, new SqlParameter[]{

                new SqlParameter("@dateB",DbType.DateTime){Value=dateB},
                new SqlParameter("@dateE",DbType.DateTime){Value=dateE}

            });


        }

        /// <summary>
        /// 方法：获得查询到的Access中数据集。
        /// </summary>
        /// <returns></returns>
        public List<TblAttList> GetAttList()
        {
            List<TblAttList> list = new List<TblAttList>();
            string strSQL = "SELECT USERID as uID, Name as uName, First(DEPTNAME) AS dep, 日期 as attDate, Min(TimeValue([checktime])) AS time1, " +
                            "Max(TimeValue([checktime])) AS time2, IIf(Format([time2] -[time1], 'hh')< 1,0,DateDiff('n', time1, time2)) AS timeDiff, " +
                            "IIf([timeDiff]=0,IIf([time1]<TimeValue('12:00:00'),'未签退','未签到'),'出勤') AS checkAtt, IIf([checkAtt]='出勤',Round([timeDiff]/60-1,3),8) AS attHours" +
                            " FROM tblAttSource GROUP BY tblAttSource.USERID, tblAttSource.Name, tblAttSource.日期";

            //string strSQL = "select * from tblAttList";
            SqlDataReader reader = SQLHelper.SelectAtt(strSQL);

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    TblAttList model = new TblAttList
                    {
                        UserID = reader.GetInt32(0),
                        //UName = reader["uName"].ToString(),
                        UName = MidString(reader.GetValue(1).ToString().Trim()),
                        Dep = reader["dep"].ToString(),
                        AttDate = reader.GetDateTime(3),
                        Time1 = reader.GetDateTime(4).ToLongTimeString(),
                        Time2 = reader.GetDateTime(5).ToLongTimeString(),
                        TimeDiff = reader.GetInt32(6),
                        CheckAtt = reader.GetString(7),
                        AttHours = reader.GetDouble(8)
                    };

                    list.Add(model);
                }

            }
            reader.Close();
            return list;

        }

        public List<TblAttSource> GetAttSource()
        {
            List<TblAttSource> list = new List<TblAttSource>();
            string strSQL = "SELECT * from tblAttSource";
            SqlDataReader reader = SQLHelper.SelectAtt(strSQL);

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    TblAttSource model = new TblAttSource
                    {
                        UserID = reader.GetInt32(0),
                        Name = MidString(reader.GetValue(1).ToString().Trim()),
                        Department = reader["DEPTNAME"].ToString(),
                        DateCheck = reader.GetDateTime(3),
                        CheckTime = reader.GetDateTime(4)
                    };

                    list.Add(model);
                }

            }
            reader.Close();
            return list;
        }


        /// <summary>
        /// 方法：获得通过查询得到的考勤原始数据（sql语句）
        /// </summary>
        /// <param name="dateB">开始日期</param>
        /// <param name="dateE">结束日期</param>
        /// <returns>返回TblAttsource的泛型集合</returns>
        public List<TblAttSource> GetAttsouceQuery(DateTime dateB, DateTime dateE)
        {
            List<TblAttSource> list = new List<TblAttSource>();
            //strSQL语句按照TeamIndex（也就是Street）、岗位进行排序。
            string strSQL = "SELECT checkinout.USERID,USERINFO.SSN,USERINFO.Name, USERINFO.CITY as DEPTNAME, CONVERT(VARCHAR(10),CHECKTIME,120) AS 日期, checkinout.checktime " +
                            "FROM(USERINFO INNER JOIN checkinout ON USERINFO.USERID = checkinout.USERID) " +
                            "WHERE(checkinout.checktime >= @dateB And checkinout.checktime < @dateE) " +
                            "order by USERINFO.STREET,USERINFO.TITLE,USERINFO.BADGENUMBER";

            SqlDataReader reader = SQLHelper.SelectAtt(strSQL, new SqlParameter[] {
                new SqlParameter("@dateB",DbType.DateTime){Value=dateB},
                new SqlParameter("@dateE",DbType.DateTime){Value=dateE}
            });

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    TblAttSource model = new TblAttSource
                    {
                        UserID = reader.GetInt32(0),
                        SSN = reader["SSN"].ToString(),
                        Name = reader.GetValue(2).ToString().Trim(),
                        Department = reader["DEPTNAME"].ToString(),
                        DateCheck = Convert.ToDateTime(reader["日期"].ToString()),
                        CheckTime = reader.GetDateTime(5)
                    };

                    list.Add(model);
                }

            }
            reader.Close();
            return list;

        }


        /// <summary>
        /// 方法：获得当前日期的出勤人数。
        /// </summary>
        /// <param name="dateE">当前日期</param>
        /// <returns></returns>
        internal object CountDateEmps(DateTime dateE)
        {
            string strSql = "select count(distinct(USERID)) as countEmps from CHECKINOUT where cast(CHECKTIME as date) = @dateE";
            return SQLHelper.ExecuteScalar(strSql, new SqlParameter[] {
            new SqlParameter("@dateE",DbType.DateTime){Value=dateE}
            });
        }

        private string MidString(string str)
        {
            string newStr = "";
            if (str != null)
            {
                if (str.Length > 5)
                {
                    newStr = str.Substring(0, 3);
                }
                else if (str.Length >= 4)
                {
                    newStr = str.Substring(0, 2);
                }
                else
                {
                    newStr = str;
                }
            }

            return newStr;
        }
    }
}
