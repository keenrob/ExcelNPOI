using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF;
using NPOI;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Collections;
using NPOI.HSSF.UserModel;

namespace ExcelNPOI
{
    public static partial class FuncExcel
    {
        /// <summary>
        /// 方法：ACCESS循环导出各车间的考勤数据到Excel中。
        /// </summary>
        /// <param name="dtBegin">周期起始</param>
        /// <param name="dtEnd">周期结束</param>
        /// <param name="wbName">文件名称</param>
        /// <param name="listAtt">规范的考勤数据</param>
        /// <param name="listAttSource">考勤原始数据</param>
        public static void CreateBook(DateTime dtBegin, DateTime dtEnd, string wbName, List<TblAttList> listAtt, List<TblAttSource> listAttSource)
        {

            IWorkbook workbook = new XSSFWorkbook();  //创建xlsx文件。
            ISheet sheet = workbook.CreateSheet(dtEnd.ToString("Y") + "考勤月报"); //创建X月报表。
            ISheet sheetSource = workbook.CreateSheet("原始数据");
            int cellCount = DateDiff(dtBegin, dtEnd) + 8;

            //打印设置
            sheet.PrintSetup.Landscape = true;//横向打印
            sheet.RepeatingRows = new CellRangeAddress(0, 2, 0, 0);//标题行
            sheet.SetMargin(MarginType.RightMargin, (double)0.1);
            sheet.SetMargin(MarginType.TopMargin, (double)0.1);
            sheet.SetMargin(MarginType.LeftMargin, (double)0.1);
            sheet.SetMargin(MarginType.BottomMargin, (double)0.1);
            //sheet.FitToPage = true;

            //顶端周期行。
            IRow rowHeader = sheet.CreateRow(0);
            rowHeader.CreateCell(0).SetCellValue("考勤周期：" + dtBegin.ToShortDateString().ToString() + "至" + dtEnd.ToShortDateString().ToString());
            //创建标题行
            IRow rowTitle1 = sheet.CreateRow(1);
            IRow rowTitle2 = sheet.CreateRow(2);
            //创建标题行数据单元格:序号、EID、姓名
            string[] strTitle1 = { "序号", "EID", "姓名" };
            for (int i = 0; i <= 2; i++) //0 1 2
            {
                CellRangeAddress region = new CellRangeAddress(1, 2, i, i);
                sheet.AddMergedRegion(region);//合并单元格。
                sheet.SetColumnWidth(i, 6 * 256);
                ICell cellTitle = rowTitle1.CreateCell(i);
                cellTitle.SetCellValue(strTitle1[i]);
                cellTitle.CellStyle = CreateTitleStyle(workbook);
                RegionUtil.SetBorderLeft(1, region, sheet, workbook);//给合并单元格画框线。

            }
            //创建日期天数标题行及数据单元格
            Dictionary<int, int> dic = new Dictionary<int, int>(); //创建一个集合用来容纳数据行的cell值。

            for (int i = 3; i <= cellCount - 5; i++)
            {
                sheet.SetColumnWidth(i, 7 * 256);

                ICell cellTitle1 = rowTitle1.CreateCell(i);
                cellTitle1.SetCellValue(Convert.ToInt32(dtBegin.AddDays(i - 3).ToString("dd")));
                cellTitle1.CellStyle = CreateTitleStyle(workbook);

                ICell cellTitle2 = rowTitle2.CreateCell(i);
                cellTitle2.SetCellValue(dtBegin.AddDays(i - 3).ToString("ddd"));
                cellTitle2.CellStyle = CreateTitleStyle(workbook);

                dic.Add(Convert.ToInt32(dtBegin.AddDays(i - 3).ToString("dd")), i); //给字典赋值：key 是日期Day；value是cell的索引号。

            }
            //创建列尾的标题行数据单元格
            string[] strTitle3 = { "出勤小时数", "未签退次数", "未签到次数", "实到天数", "夜餐次数" };
            for (int i = 4; i >= 0; i--)
            {
                CellRangeAddress region = new CellRangeAddress(1, 2, cellCount - i, cellCount - i);
                sheet.AddMergedRegion(region);
                ICell cellTitle3 = rowTitle1.CreateCell(cellCount - i);
                cellTitle3.SetCellValue(strTitle3[i]);
                cellTitle3.CellStyle = CreateTitleStyle(workbook);
                RegionUtil.SetBorderRight(1, region, sheet, workbook);//给合并单元格画框线。
            }

            //循环写入数据
            IRow rowData;
            ICell cellData;
            List<int> listUser = listAtt.Select(u => u.UserID).Distinct().ToList(); //获得集合中的UserID集合。
            List<TblAttList> lists = null;

            //先创建单元格样式，并赋值对应的样式内容。注意：不能将创建样式的内容写到大循环里面，会导致内存溢出。这个问题居然困惑我一个礼拜了。o(╥﹏╥)o
            ICellStyle cellStyleContent = workbook.CreateCellStyle();
            cellStyleContent = CreateContentStyle(workbook);

            ICellStyle cellStyleTitle = workbook.CreateCellStyle();
            cellStyleTitle = CreateTitleStyle(workbook);

            //根据UserID再筛选出相关的数据集合
            //listAtt.Where(u => u.UserID == listUser[0]).ToList();
            if (listUser.Count() > 0) //集合数据集大于0
            {
                for (int i = 0; i < listUser.Count; i++)
                {
                    //根据UserID再筛选出相关的数据集合
                    lists = listAtt.Where(u => u.UserID == listUser[i]).ToList();

                    rowData = sheet.CreateRow(i + 3);//根据员工的数据建立行。从第4行开始；
                    rowData.Height = 23 * 20;

                    //建立序号、EID、姓名单元格 即左侧 Left
                    string[] strLeft = { Convert.ToString(i + 1), lists[0].UserID.ToString(), lists[0].UName };
                    //var strLefts = new { intNumber = i + 1, userID = lists[0].UserID, name = lists[0].UName };//创建匿名类。

                    for (int j = 0; j < 3; j++)
                    {
                        ICell cellLeft = rowData.CreateCell(j);
                        cellLeft.SetCellValue(strLeft[j]);
                        cellLeft.CellStyle = cellStyleContent;
                    }

                    int checkAttIn = 0;
                    int checkAttOut = 0;
                    double attHours = 0.00;
                    int dinnerCount = 0;
                    List<int> attIntList = new List<int>(); //建立一个int集合，长度为该员工的所有考勤数据条数。准备用考勤数据的日期（day值）填充。
                    //遍历lists里面某员工的所有考勤数据。
                    foreach (TblAttList item in lists)
                    {
                        int _day = Convert.ToInt32(item.AttDate.ToString("dd"));
                        attIntList.Add(_day);
                        cellData = rowData.CreateCell(dic[_day]); //日期天所对应的value值。
                        string time1 = item.CheckAtt == "未签到" ? "-" : item.Time1;
                        string time2 = item.CheckAtt == "未签退" ? "-" : item.Time2;

                        cellData.SetCellValue(time1 + "\n" + time2);
                        // cellData.SetCellValue(item.AttDate.ToShortDateString().ToString());//显示日期。对应。


                        //cellData.CellStyle = CreateContentStyle(workbook); //你就等着溢出吧 O(∩_∩)O哈哈~

                        cellData.CellStyle = cellStyleContent;


                        if (item.CheckAtt == "未签到")
                        {
                            checkAttIn += 1;
                        }

                        checkAttOut = item.CheckAtt == "未签退" ? checkAttOut + 1 : checkAttOut;
                        attHours += item.AttHours;

                        if (time2 != "-")
                        {
                            //10:00:00,09:00:00  9:00:00

                            if (time2.Length == 8 && int.Parse(time2.Substring(0, 2)) >= 22)
                            {
                                dinnerCount += 1;
                            }
                        }

                    }

                    int[] dayInts = dic.Keys.ToArray<int>(); //获得键值对集合中，键的集合。
                    int[] en2 = dayInts.Concat(attIntList).Except(dayInts.Intersect(attIntList)).ToArray();// 容斥原理
                    for (int s = 0; s < en2.Count(); s++)
                    {
                        cellData = rowData.CreateCell(dic[en2[s]]); //日期天所对应的value值。
                        cellData.SetCellValue("箜");
                        cellData.CellStyle = cellStyleContent;
                    }

                    #region 建立实到天数等单元格，即最右侧 Right
                    ICell cellRight0 = rowData.CreateCell(cellCount - 4, CellType.Numeric);  //晚餐判断次数>=22点
                    if (dinnerCount == 0)
                    {
                        cellRight0.SetCellValue("");
                    }
                    else
                    {
                        cellRight0.SetCellValue(dinnerCount);
                    }
                    cellRight0.CellStyle = cellStyleTitle;

                    ICell cellRight1 = rowData.CreateCell(cellCount - 3);//实到天数
                    cellRight1.SetCellValue(lists.Count());
                    cellRight1.CellStyle = cellStyleTitle;

                    ICell cellRight2 = rowData.CreateCell(cellCount - 2);//未签到次数
                    if (checkAttIn == 0)
                    {
                        cellRight2.SetCellValue("");
                    }
                    else
                    {
                        cellRight2.SetCellValue(checkAttIn);
                    }
                    cellRight2.CellStyle = cellStyleTitle;

                    ICell cellRight3 = rowData.CreateCell(cellCount - 1);//未签退次数
                    if (checkAttOut == 0)
                    {
                        cellRight3.SetCellValue("");
                    }
                    else
                    {
                        cellRight3.SetCellValue(checkAttOut);
                    }
                    cellRight3.CellStyle = cellStyleTitle;

                    ICell cellRight4 = rowData.CreateCell(cellCount);//累计小时数
                    cellRight4.SetCellValue(attHours);
                    cellRight4.CellStyle = cellStyleTitle;
                    #endregion




                }
            }

            //创建原始数据表
            IRow rowSourceTitle = sheetSource.CreateRow(0);
            string[] strSourceTitle = { "员工ID", "姓名", "部门", "日期", "打卡时间" };
            for (int i = 0; i <= 4; i++) //0 1 2 3 4
            {
                sheetSource.SetColumnWidth(i, 15 * 256);
                ICell cellTitle = rowSourceTitle.CreateCell(i);
                cellTitle.SetCellValue(strSourceTitle[i]);

            }
            IRow rowSourceContent;

            for (int i = 1; i < listAttSource.Count; i++)
            {
                rowSourceContent = sheetSource.CreateRow(i);
                rowSourceContent.CreateCell(0).SetCellValue(listAttSource[i - 1].UserID);
                rowSourceContent.CreateCell(1).SetCellValue(listAttSource[i - 1].Name);
                rowSourceContent.CreateCell(2).SetCellValue(listAttSource[i - 1].Department);
                rowSourceContent.CreateCell(3).SetCellValue(listAttSource[i - 1].DateCheck.ToShortDateString());
                rowSourceContent.CreateCell(4).SetCellValue(listAttSource[i - 1].CheckTime.ToLongTimeString());


            }


            //写入文件：

            try
            {
                if (File.Exists(wbName))
                {
                    File.Delete(wbName);
                }
            }
            catch (Exception)
            {

                throw;

            }



            using (FileStream fs = new FileStream(wbName, FileMode.OpenOrCreate))
            {
                workbook.Write(fs);
                workbook.Close();
                workbook = null;
            }



        }

        /// <summary>
        /// 方法：循环导出各车间的考勤数据到Excel中
        /// </summary>
        /// <param name="dtBegin"></param>
        /// <param name="dtEnd"></param>
        /// <param name="wbName"></param>
        /// <param name="listAttSource"></param>
        public static void CreateBookForCJ(DateTime dtBegin, DateTime dtEnd, string wbName, List<TblAttSource> listAttSource)
        {

            IWorkbook workbook = new XSSFWorkbook();  //创建xlsx文件。
            ISheet sheet = workbook.CreateSheet(dtEnd.ToString("Y") + "考勤月报"); //创建X月报表。
            ISheet sheetSource = workbook.CreateSheet("原始数据");
            int cellCount = DateDiff(dtBegin, dtEnd) + 10;

            //打印设置
            sheet.PrintSetup.Landscape = true;//横向打印
            sheet.RepeatingRows = new CellRangeAddress(0, 2, 0, 0);//标题行
            sheet.SetMargin(MarginType.RightMargin, (double)0.1);
            sheet.SetMargin(MarginType.TopMargin, (double)0.1);
            sheet.SetMargin(MarginType.LeftMargin, (double)0.1);
            sheet.SetMargin(MarginType.BottomMargin, (double)0.1);
            //sheet.FitToPage = true;

            //顶端周期行。
            IRow rowHeader = sheet.CreateRow(0);
            rowHeader.CreateCell(0).SetCellValue("考勤周期：" + dtBegin.ToShortDateString().ToString() + "至" + dtEnd.ToShortDateString().ToString());
            //创建标题行
            IRow rowTitle1 = sheet.CreateRow(1);
            IRow rowTitle2 = sheet.CreateRow(2);
            //创建标题行数据单元格:序号、EID、姓名
            string[] strTitle1 = { "序号", "eID", "eNO","姓名" };
            for (int i = 0; i <= 3; i++) //0 1 2
            {
                CellRangeAddress region = new CellRangeAddress(1, 2, i, i);
                sheet.AddMergedRegion(region);//合并单元格。
                sheet.SetColumnWidth(i, 6 * 256);
                ICell cellTitle = rowTitle1.CreateCell(i);
                cellTitle.SetCellValue(strTitle1[i]);
                cellTitle.CellStyle = CreateTitleStyle(workbook);
                RegionUtil.SetBorderLeft(1, region, sheet, workbook);//给合并单元格画框线。

            }
            //创建日期天数标题行及数据单元格
            Dictionary<int, int> dic = new Dictionary<int, int>(); //创建一个集合用来容纳数据行的cell值。

            for (int i = 4; i <= cellCount - 6; i++)
            {
                sheet.SetColumnWidth(i, 7 * 256);

                ICell cellTitle1 = rowTitle1.CreateCell(i);
                cellTitle1.SetCellValue(Convert.ToInt32(dtBegin.AddDays(i - 4).ToString("dd")));
                cellTitle1.CellStyle = CreateTitleStyle(workbook);

                ICell cellTitle2 = rowTitle2.CreateCell(i);
                cellTitle2.SetCellValue(dtBegin.AddDays(i - 4).ToString("ddd"));
                cellTitle2.CellStyle = CreateTitleStyle(workbook);

                dic.Add(Convert.ToInt32(dtBegin.AddDays(i - 4).ToString("dd")), i); //给字典赋值：key 是日期Day；value是cell的索引号。

            }
            //创建列尾的标题行数据单元格
            string[] strTitle3 = { "折算工数", "出勤小时数", "未签退次数", "未签到次数", "实到天数", "夜餐次数" };
            for (int i = 5; i >= 0; i--)
            {
                CellRangeAddress region = new CellRangeAddress(1, 2, cellCount - i, cellCount - i);
                sheet.AddMergedRegion(region);
                ICell cellTitle3 = rowTitle1.CreateCell(cellCount - i);
                cellTitle3.SetCellValue(strTitle3[i]);
                cellTitle3.CellStyle = CreateTitleStyle(workbook);
                RegionUtil.SetBorderRight(1, region, sheet, workbook);//给合并单元格画框线。
            }

            //Linq方式获得新的list集合
            var listAtt = from a in listAttSource
                        group a by new { a.UserID, a.DateCheck } into m
                        // where m.Key.UserID == 872
                        select new
                        {
                            m.Key.UserID,
                            m.FirstOrDefault().SSN,
                            AttDate = m.Key.DateCheck,
                            UName = m.FirstOrDefault().Name,
                            department = m.FirstOrDefault().Department,
                            checkAm = TimeCheckAm(m.Min(am => am.CheckTime)),
                            checkPm = TimeCheckPm(m.Max(pm => pm.CheckTime))

                        };

            //循环写入数据
            IRow rowData;
            ICell cellData;
            List<int> listUser = listAtt.Select(u => u.UserID).Distinct().ToList(); //获得集合中的UserID集合。
            //List<TblAttList> lists = null;

            //先创建单元格样式，并赋值对应的样式内容。注意：不能将创建样式的内容写到大循环里面，会导致内存溢出。这个问题居然困惑我一个礼拜了。o(╥﹏╥)o
            ICellStyle cellStyleContent = workbook.CreateCellStyle();
            cellStyleContent = CreateContentStyle(workbook);

            ICellStyle cellStyleTitle = workbook.CreateCellStyle();
            cellStyleTitle = CreateTitleStyle(workbook);

            //根据UserID再筛选出相关的数据集合
            //listAtt.Where(u => u.UserID == listUser[0]).ToList();
            if (listUser.Count() > 0) //集合数据集大于0
            {
                for (int i = 0; i < listUser.Count; i++)
                {
                    //根据UserID再筛选出相关的数据集合
                    var lists = listAtt.Where(u => u.UserID == listUser[i]).ToList();

                    rowData = sheet.CreateRow(i + 3);//根据员工的数据建立行。从第4行开始；
                    rowData.Height = 30 * 20;

                    //建立序号、EID、姓名单元格 即左侧 Left
                    string[] strLeft = { Convert.ToString(i + 1), lists[0].UserID.ToString(), lists[0].SSN,lists[0].UName };

                    for (int j = 0; j < 4; j++)
                    {
                        ICell cellLeft = rowData.CreateCell(j);
                        cellLeft.SetCellValue(strLeft[j]);
                        cellLeft.CellStyle = cellStyleContent;
                    }

                    int dinnerCount = 0;//晚餐次数：>=22:00
                    int checkAttIn = 0; //未签到次数
                    int checkAttOut = 0;//未签退次数
                    double attHours = 0.00;//合计出勤小时数量
                    double attDays = 0.00;//折算天数

                    List<int> attIntList = new List<int>(); //建立一个int集合，长度为该员工的所有考勤数据条数。准备用考勤数据的日期（day值）填充。
                    //遍历lists里面某员工的所有考勤数据。
                    foreach (var item in lists)
                    {
                        int _day = Convert.ToInt32(item.AttDate.ToString("dd"));
                        attIntList.Add(_day);
                        cellData = rowData.CreateCell(dic[_day]); //日期天所对应的value值。
                        string time1 = item.checkAm.Hour == 0 ? "-" : item.checkAm.ToString("HH:mm:ss");
                        string time2 = item.checkPm.Hour == 0 ? "-" : item.checkPm.ToString("HH:mm:ss");

                        cellData.SetCellValue(time1 + "\n" + time2);
                        // cellData.SetCellValue(item.AttDate.ToShortDateString().ToString());//显示日期。对应。
                        //cellData.CellStyle = CreateContentStyle(workbook); //你就等着溢出吧 O(∩_∩)O哈哈~

                        cellData.CellStyle = cellStyleContent;

                        //未签到次数
                        if (time1 == "-")
                        {
                            checkAttIn += 1;
                        }

                        //未签退次数：
                        checkAttOut = time2 == "-" ? checkAttOut + 1 : checkAttOut;

                        //有两种情况，如果未签到，则默认为8:00签到；如果未签退，则默认17:00签退。
                        DateTime amCheckNew = time1 == "-" ? Convert.ToDateTime(item.checkPm.ToShortDateString()).AddHours(8).AddMinutes(0) : item.checkAm;
                        DateTime pmCheckNew = time2 == "-" ? Convert.ToDateTime(item.checkPm.ToShortDateString()).AddHours(17).AddMinutes(0) : item.checkPm;

                        //日工作时长：如果工作时长大于等于5小时，则减掉1小时午饭时间。用floor的目的是减少一些分钟的误差。
                        TimeSpan sp = pmCheckNew - amCheckNew;
                        attHours += Math.Floor((Math.Round(sp.TotalMinutes / 60, 2) >= 5 ? Math.Round(sp.TotalMinutes / 60, 2) - 1 : Math.Round(sp.TotalMinutes / 60, 2))*10)/10;

                        attDays =Math.Floor(Math.Round(attHours / 8, 2)*10)/10; //保留1位小数但不四舍五入。


                        //晚餐补贴次数。
                        if (time2 != "-")
                        {
                            //10:00:00,09:00:00  9:00:00

                            if (time2.Length == 8 && int.Parse(time2.Substring(0, 2)) >= 22)
                            {
                                dinnerCount += 1;
                            }
                        }

                    }

                    int[] dayInts = dic.Keys.ToArray<int>(); //获得键值对集合中，键的集合。
                    int[] en2 = dayInts.Concat(attIntList).Except(dayInts.Intersect(attIntList)).ToArray();// 容斥原理
                    for (int s = 0; s < en2.Count(); s++)
                    {
                        cellData = rowData.CreateCell(dic[en2[s]]); //日期天所对应的value值。
                        cellData.SetCellValue("箜");
                        cellData.CellStyle = cellStyleContent;
                    }

                    #region 建立实到天数等单元格，即最右侧 Right
                    ICell cellRight0 = rowData.CreateCell(cellCount - 5, CellType.Numeric);  //晚餐判断次数>=22点
                    if (dinnerCount == 0)
                    {
                        cellRight0.SetCellValue("");
                    }
                    else
                    {
                        cellRight0.SetCellValue(dinnerCount);
                    }
                    cellRight0.CellStyle = cellStyleTitle;

                    ICell cellRight1 = rowData.CreateCell(cellCount - 4);//实到天数
                    cellRight1.SetCellValue(lists.Count());
                    cellRight1.CellStyle = cellStyleTitle;

                    ICell cellRight2 = rowData.CreateCell(cellCount - 3);//未签到次数
                    if (checkAttIn == 0)
                    {
                        cellRight2.SetCellValue("");
                    }
                    else
                    {
                        cellRight2.SetCellValue(checkAttIn);
                    }
                    cellRight2.CellStyle = cellStyleTitle;

                    ICell cellRight3 = rowData.CreateCell(cellCount - 2);//未签退次数
                    if (checkAttOut == 0)
                    {
                        cellRight3.SetCellValue("");
                    }
                    else
                    {
                        cellRight3.SetCellValue(checkAttOut);
                    }
                    cellRight3.CellStyle = cellStyleTitle;

                    ICell cellRight4 = rowData.CreateCell(cellCount-1);//累计小时数
                    cellRight4.SetCellValue(attHours);
                    cellRight4.CellStyle = cellStyleTitle;

                    ICell cellRight5 = rowData.CreateCell(cellCount);//折算天数
                    cellRight5.SetCellValue(attDays);
                    cellRight5.CellStyle = cellStyleTitle;

                    #endregion




                }
            }

            //创建原始数据表
            IRow rowSourceTitle = sheetSource.CreateRow(0);
            string[] strSourceTitle = { "员工ID", "姓名", "部门", "日期", "打卡时间" };
            for (int i = 0; i <= 4; i++) //0 1 2 3 4
            {
                sheetSource.SetColumnWidth(i, 15 * 256);
                ICell cellTitle = rowSourceTitle.CreateCell(i);
                cellTitle.SetCellValue(strSourceTitle[i]);

            }
            IRow rowSourceContent;

            for (int i = 1; i < listAttSource.Count; i++)
            {
                rowSourceContent = sheetSource.CreateRow(i);
                rowSourceContent.CreateCell(0).SetCellValue(listAttSource[i - 1].UserID);
                rowSourceContent.CreateCell(1).SetCellValue(listAttSource[i - 1].Name);
                rowSourceContent.CreateCell(2).SetCellValue(listAttSource[i - 1].Department);
                rowSourceContent.CreateCell(3).SetCellValue(listAttSource[i - 1].DateCheck.ToShortDateString());
                rowSourceContent.CreateCell(4).SetCellValue(listAttSource[i - 1].CheckTime.ToLongTimeString());


            }


            //写入文件：

            try
            {
                if (File.Exists(wbName))
                {
                    File.Delete(wbName);
                }
            }
            catch (Exception)
            {

                throw;

            }



            using (FileStream fs = new FileStream(wbName, FileMode.OpenOrCreate))
            {
                workbook.Write(fs);
                workbook.Close();
                workbook = null;
            }



        }


        /// <summary>
        /// 方法：导出行政各部门的考勤数据到Excel中。
        /// </summary>
        /// <param name="dtBegin">周期起始</param>
        /// <param name="dtEnd">周期结束</param>
        /// <param name="wbName">文件名称</param>
        /// <param name="listAtt">规范的考勤数据</param>
        /// <param name="listAttSource">考勤原始数据</param>
        public static void CreateBookForManage(DateTime dtBegin, DateTime dtEnd, string wbName, List<TblAttSource> listAttSource)
        {

            IWorkbook workbook = new XSSFWorkbook();  //创建xlsx文件。
            ISheet sheet = workbook.CreateSheet(dtEnd.ToString("Y") + "考勤月报"); //创建X月报表。
            ISheet sheetSource = workbook.CreateSheet("原始数据");
            int cellCount = DateDiff(dtBegin, dtEnd) + 11; //周期差值+额外的10个标题列

            //打印设置
            sheet.PrintSetup.Landscape = true;//横向打印
            sheet.RepeatingRows = new CellRangeAddress(0, 2, 0, 0);//标题行
            sheet.SetMargin(MarginType.RightMargin, (double)0.1);
            sheet.SetMargin(MarginType.TopMargin, (double)0.1);
            sheet.SetMargin(MarginType.LeftMargin, (double)0.1);
            sheet.SetMargin(MarginType.BottomMargin, (double)0.1);
            //sheet.FitToPage = true;



            //顶端周期行。
            IRow rowHeader = sheet.CreateRow(0);
            rowHeader.CreateCell(0).SetCellValue("考勤周期：" + dtBegin.ToShortDateString().ToString() + "至" + dtEnd.ToShortDateString().ToString());
            //创建标题行
            IRow rowTitle1 = sheet.CreateRow(1);
            IRow rowTitle2 = sheet.CreateRow(2);
            //创建标题行数据单元格:序号、EID、姓名
            string[] strTitle1 = { "序号", "eID","eNO","姓名", "部门" };
            for (int i = 0; i <= 4; i++) //0 1 2 3 4
            {
                CellRangeAddress region = new CellRangeAddress(1, 2, i, i);
                sheet.AddMergedRegion(region);//合并单元格。
                sheet.SetColumnWidth(i, 6 * 256);
                ICell cellTitle = rowTitle1.CreateCell(i);
                cellTitle.SetCellValue(strTitle1[i]);
                cellTitle.CellStyle = CreateTitleStyle(workbook);
                RegionUtil.SetBorderLeft(1, region, sheet, workbook);//给合并单元格画框线。

            }
            //创建日期天数标题行及数据单元格
            Dictionary<int, int> dic = new Dictionary<int, int>(); //创建一个集合用来容纳数据行的cell值。

            for (int i = 5; i <= cellCount - 6; i++)
            {
                sheet.SetColumnWidth(i, 7 * 256);

                ICell cellTitle1 = rowTitle1.CreateCell(i);
                cellTitle1.SetCellValue(Convert.ToInt32(dtBegin.AddDays(i - 5).ToString("dd")));
                cellTitle1.CellStyle = CreateTitleStyle(workbook);

                ICell cellTitle2 = rowTitle2.CreateCell(i);
                cellTitle2.SetCellValue(dtBegin.AddDays(i - 5).ToString("ddd"));
                cellTitle2.CellStyle = CreateTitleStyle(workbook);

                dic.Add(Convert.ToInt32(dtBegin.AddDays(i - 5).ToString("dd")), i); //给字典赋值：key 是日期Day；value是cell的索引号。

            }
            //创建列尾的标题行数据单元格
            string[] strTitle3 = { "出勤小时数", "未签退次数", "未签到次数", "早退次数", "迟到次数", "实到天数" };
            for (int i = 5; i >= 0; i--)
            {
                CellRangeAddress region = new CellRangeAddress(1, 2, cellCount - i, cellCount - i);
                sheet.AddMergedRegion(region);
                sheet.SetColumnWidth(cellCount - i, 6 * 256);
                ICell cellTitle3 = rowTitle1.CreateCell(cellCount - i);
                cellTitle3.SetCellValue(strTitle3[i]);
                cellTitle3.CellStyle = CreateTitleStyle(workbook);
                RegionUtil.SetBorderRight(1, region, sheet, workbook);//给合并单元格画框线。

            }

            //Linq方式获得新的list集合
            var query = from a in listAttSource
                        group a by new { a.UserID, a.DateCheck } into m
                        // where m.Key.UserID == 872
                        select new
                        {
                            m.Key.UserID,
                            m.FirstOrDefault().SSN,
                            AttDate = m.Key.DateCheck,
                            UName = m.FirstOrDefault().Name,
                            department = m.FirstOrDefault().Department,
                            checkAm = TimeCheckAm(m.Min(am => am.CheckTime), out string strAmLate),
                            checkMm = TimeCheckMm(m.Select(Mm => Mm.CheckTime).ToList(), out string strMmLate),
                            checkPm = TimeCheckPm(m.Max(pm => pm.CheckTime), out string strLE),
                            checkAmLate = strAmLate,
                            checkMmLate = strMmLate,
                            checkPmLe = strLE


                        };
            var listAtt = query.OrderBy(q => q.UserID).OrderBy(q => q.department).ToList();

            //循环写入数据
            IRow rowData;
            ICell cellData;
            List<int> listUser = listAtt.Select(u => u.UserID).Distinct().ToList(); //获得集合中的UserID集合。


            //先创建单元格样式，并赋值对应的样式内容。注意：不能将创建样式的内容写到大循环里面，会导致内存溢出。这个问题居然困惑我一个礼拜了。o(╥﹏╥)o
            ICellStyle cellStyleContent = workbook.CreateCellStyle();  //普通内容样式
            cellStyleContent = CreateContentStyle(workbook);

            ICellStyle cellStyleContentRed = workbook.CreateCellStyle();  //红色内容样式
            cellStyleContentRed = CreateContentStyleRed(workbook);

            ICellStyle cellStyleTitle = workbook.CreateCellStyle();
            cellStyleTitle = CreateTitleStyle(workbook);

            //根据UserID再筛选出相关的数据集合
            //listAtt.Where(u => u.UserID == listUser[0]).ToList();
            if (listUser.Count() > 0) //集合数据集大于0
            {
                for (int i = 0; i < listUser.Count; i++)
                {
                    //根据UserID再筛选出相关的数据集合
                    var lists = listAtt.Where(u => u.UserID == listUser[i]).ToList();

                    rowData = sheet.CreateRow(i + 3);//根据员工的数据建立行。从第4行开始；
                    rowData.Height = 30 * 20;

                    //建立序号、EID、姓名、部门单元格 即左侧 Left
                    string[] strLeft = { Convert.ToString(i + 1), lists[0].UserID.ToString(),lists[0].SSN, lists[0].UName, lists[0].department };
                    //var strLefts = new { intNumber = i + 1, userID = lists[0].UserID, name = lists[0].UName };//创建匿名类。

                    for (int j = 0; j < 5; j++)
                    {
                        ICell cellLeft = rowData.CreateCell(j);
                        cellLeft.SetCellValue(strLeft[j]);
                        cellLeft.CellStyle = cellStyleContent;
                    }

                    int checkAttIn = 0; //未签到次数
                    int checkAttNoonIn = 0;//中午未签到次数
                    int checkAttOut = 0;//未签退次数
                    int checkAmLate = 0;//上午迟到次数
                    int checkMmLate = 0;//中午迟到次数
                    int checkLE = 0;//下午早退次数
                    double attHours = 0.00;
                    List<int> attIntList = new List<int>(); //建立一个int集合，长度为该员工的所有考勤数据条数。准备用考勤数据的日期（day值）填充。
                    //遍历lists里面某员工的所有考勤数据。
                    foreach (var item in lists)
                    {
                        int _day = Convert.ToInt32(item.AttDate.ToString("dd"));
                        attIntList.Add(_day);
                        cellData = rowData.CreateCell(dic[_day]); //日期天所对应的value值。
                        string time1 = item.checkAm.Hour == 0 ? "-" : item.checkAm.ToString("HH:mm:ss");
                        string time2 = item.checkMm.Hour == 0 ? "-" : item.checkMm.ToString("HH:mm:ss");
                        string time3 = item.checkPm.Hour == 0 ? "-" : item.checkPm.ToString("HH:mm:ss");

                        cellData.SetCellValue(time1 + "\n" + time2 + "\n" + time3);
                        // cellData.SetCellValue(item.AttDate.ToShortDateString().ToString());//显示日期。对应。


                        //cellData.CellStyle = CreateContentStyle(workbook); //你就等着溢出吧 O(∩_∩)O哈哈~

                        if (item.AttDate.DayOfWeek.ToString() == "Saturday" || item.AttDate.DayOfWeek.ToString() == "Sunday") //如果当前日期为周六或周日，则将数据设置为红色。
                        {
                            cellData.CellStyle = cellStyleContentRed;

                        }
                        else
                        {
                            cellData.CellStyle = cellStyleContent;
                        }

                        //早未签到统计
                        if (item.checkAm.Hour == 0)
                        {
                            checkAttIn += 1;
                        }
                        //中午未签到统计
                        checkAttNoonIn = item.checkMm.Hour == 0 ? (item.AttDate.DayOfWeek.ToString() == "Saturday" || item.AttDate.DayOfWeek.ToString() == "Sunday") ? checkAttNoonIn : checkAttNoonIn + 1 : checkAttNoonIn;
                        //下午未签退统计
                        checkAttOut = item.checkPm.Hour == 0 ? checkAttOut + 1 : checkAttOut;
                        //早上迟到统计
                        checkAmLate = item.checkAmLate == "迟到" ? (item.AttDate.DayOfWeek.ToString() == "Saturday" || item.AttDate.DayOfWeek.ToString() == "Sunday") ? checkAmLate : checkAmLate + 1 : checkAmLate;
                        //中午迟到统计
                        checkMmLate = item.checkMmLate == "迟到" ? (item.AttDate.DayOfWeek.ToString() == "Saturday" || item.AttDate.DayOfWeek.ToString() == "Sunday") ? checkMmLate : checkMmLate + 1 : checkMmLate;
                        //下午早退统计
                        checkLE = item.checkPmLe == "早退" ? (item.AttDate.DayOfWeek.ToString() == "Saturday" || item.AttDate.DayOfWeek.ToString() == "Sunday") ? checkLE : checkLE + 1 : checkLE;
                        //attHours += item.AttHours;

                        //有两种情况，如果未签到，则默认为8:00签到；如果未签退，则默认17:00签退。
                        DateTime amCheckNew = time1 == "-" ? Convert.ToDateTime(item.checkPm.ToShortDateString()).AddHours(8).AddMinutes(0) : item.checkAm;
                        DateTime MmCheckNew = item.checkMm;
                        DateTime pmCheckNew = time3 == "-" ? Convert.ToDateTime(item.checkPm.ToShortDateString()).AddHours(17).AddMinutes(0) : item.checkPm;

                        //日工作时长：如果工作时长大于等于5小时，则减掉1小时午饭时间。用floor的目的是减少一些分钟的误差。
                        //如果中午没有刷卡，则下午-上午，如果刷卡了，则中午-上午 加上 下午-中午。
                        TimeSpan sp = MmCheckNew.Hour == 0 ? pmCheckNew - amCheckNew : (MmCheckNew - amCheckNew)+(pmCheckNew-MmCheckNew);
                        attHours += Math.Floor((Math.Round(sp.TotalMinutes / 60, 2) >= 5 ? Math.Round(sp.TotalMinutes / 60, 2) - 1 : Math.Round(sp.TotalMinutes / 60, 2)) * 10) / 10;

                        //attDays = Math.Floor(Math.Round(attHours / 8, 2) * 10) / 10; //保留1位小数但不四舍五入。


                    }

                    int[] dayInts = dic.Keys.ToArray<int>(); //获得键值对集合中，键的集合。
                    int[] en2 = dayInts.Concat(attIntList).Except(dayInts.Intersect(attIntList)).ToArray();// 容斥原理
                    for (int s = 0; s < en2.Count(); s++)
                    {
                        cellData = rowData.CreateCell(dic[en2[s]]); //日期天所对应的value值。
                        cellData.SetCellValue("箜");
                        cellData.CellStyle = cellStyleContent;

                    }


                    //建立实到天数等单元格，即最右侧 Right
                    ICell cellRight1 = rowData.CreateCell(cellCount - 5);
                    cellRight1.SetCellValue(lists.Count());
                    cellRight1.CellStyle = cellStyleTitle;

                    ICell cellRight2 = rowData.CreateCell(cellCount - 4);
                    cellRight2.SetCellValue("早:" + checkAmLate + "\n" + "午:" + checkMmLate);
                    //cellRight2.SetCellValue( checkAmLate);
                    cellRight2.CellStyle = cellStyleTitle;

                    ICell cellRight3 = rowData.CreateCell(cellCount - 3);
                    cellRight3.SetCellValue(checkLE);
                    cellRight3.CellStyle = cellStyleTitle;

                    ICell cellRight4 = rowData.CreateCell(cellCount - 2);
                    cellRight4.SetCellValue("早:" + checkAttIn + "\n" + "午:" + checkAttNoonIn);
                    cellRight4.CellStyle = cellStyleTitle;

                    ICell cellRight5 = rowData.CreateCell(cellCount - 1);
                    cellRight5.SetCellValue(checkAttOut);
                    cellRight5.CellStyle = cellStyleTitle;

                    ICell cellRight6 = rowData.CreateCell(cellCount);
                    cellRight6.SetCellValue(attHours);
                    cellRight6.CellStyle = cellStyleTitle;


                }
            }

            //冻结左边4列，上面3行
            sheet.CreateFreezePane(0, 3, 0, 0);   //发生显示错位。不知何故。

            //创建原始数据表
            IRow rowSourceTitle = sheetSource.CreateRow(0);
            string[] strSourceTitle = { "员工ID", "姓名", "部门", "日期", "打卡时间" };
            for (int i = 0; i <= 4; i++) //0 1 2 3 4
            {
                sheetSource.SetColumnWidth(i, 15 * 256);
                ICell cellTitle = rowSourceTitle.CreateCell(i);
                cellTitle.SetCellValue(strSourceTitle[i]);

            }
            IRow rowSourceContent;
            listAttSource = listAttSource.OrderBy(x => x.CheckTime).OrderBy(x => x.DateCheck).OrderBy(x => x.UserID).OrderBy(x => x.Department).ToList();//排序。
            for (int i = 1; i < listAttSource.Count; i++)
            {
                rowSourceContent = sheetSource.CreateRow(i);
                rowSourceContent.CreateCell(0).SetCellValue(listAttSource[i - 1].UserID);
                rowSourceContent.CreateCell(1).SetCellValue(listAttSource[i - 1].Name);
                rowSourceContent.CreateCell(2).SetCellValue(listAttSource[i - 1].Department);
                rowSourceContent.CreateCell(3).SetCellValue(listAttSource[i - 1].DateCheck.ToShortDateString());
                rowSourceContent.CreateCell(4).SetCellValue(listAttSource[i - 1].CheckTime.ToLongTimeString());


            }


            //写入文件：

            try
            {
                if (File.Exists(wbName))
                {
                    File.Delete(wbName);
                }
            }
            catch (Exception)
            {

                throw;

            }



            using (FileStream fs = new FileStream(wbName, FileMode.OpenOrCreate))
            {
                workbook.Write(fs);
                workbook.Close();
                workbook = null;
            }



        }

        /// <summary>
        /// 方法：计算两个日期差值。
        /// </summary>
        /// <param name="dateStart"></param>
        /// <param name="dateEnd"></param>
        /// <returns></returns>
        private static int DateDiff(DateTime dateStart, DateTime dateEnd)
        {
            DateTime start = Convert.ToDateTime(dateStart.ToShortDateString());
            DateTime end = Convert.ToDateTime(dateEnd.ToShortDateString());
            TimeSpan sp = end.Subtract(start);
            return sp.Days;
        }

        /// <summary>
        /// 方法：设置Excel标题行样式
        /// </summary>
        /// <param name="book"></param>
        /// <returns></returns>
        public static ICellStyle CreateTitleStyle(IWorkbook book)
        {
            ICellStyle cellStyle = book.CreateCellStyle();
            cellStyle.WrapText = true;
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;


            IFont fontLeft = book.CreateFont();
            fontLeft.FontHeightInPoints = 9;
            //fontLeft.Boldweight = (short)FontBoldWeight.Bold;
            fontLeft.FontName = "微软雅黑";
            cellStyle.ShrinkToFit = true;
            cellStyle.SetFont(fontLeft);
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            return cellStyle;
        }

        /// <summary>
        /// 方法：设置Excel数据行样式(普通样式)
        /// </summary>
        /// <param name="book"></param>
        /// <returns></returns>
        public static ICellStyle CreateContentStyle(IWorkbook book)
        {
            ICellStyle cellStyle = book.CreateCellStyle();
            cellStyle.WrapText = true;
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;


            IFont fontLeft = book.CreateFont();
            fontLeft.FontHeightInPoints = 8;
            fontLeft.FontName = "宋体";
            cellStyle.ShrinkToFit = true;
            cellStyle.SetFont(fontLeft);
            cellStyle.BorderLeft = BorderStyle.Thin;
            //cellStyle.BorderTop = BorderStyle.Thin;
            //cellStyle.BorderBottom = BorderStyle.Thin;
            return cellStyle;
        }

        /// <summary>
        /// 方法：设置Excel数据防样式（前景色为红色）
        /// </summary>
        /// <param name="book"></param>
        /// <returns></returns>
        public static ICellStyle CreateContentStyleRed(IWorkbook book)
        {
            ICellStyle cellStyle = book.CreateCellStyle();
            cellStyle.WrapText = true;
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;


            IFont fontLeft = book.CreateFont();
            fontLeft.FontHeightInPoints = 8;
            fontLeft.FontName = "宋体";
            fontLeft.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
            cellStyle.ShrinkToFit = true;
            cellStyle.SetFont(fontLeft);
            cellStyle.BorderLeft = BorderStyle.Thin;
            //cellStyle.BorderTop = BorderStyle.Thin;
            //cellStyle.BorderBottom = BorderStyle.Thin;
            return cellStyle;
        }


        /// <summary>
        /// 方法：判断早上签到时间
        /// </summary>
        /// <param name="time">打卡时间</param>
        /// <returns></returns>
        private static DateTime TimeCheckAm(DateTime time, out string strLate)
        {
            DateTime checkAM = Convert.ToDateTime(time.ToShortDateString()).AddHours(11).AddMinutes(1);
            DateTime checkAMLate = Convert.ToDateTime(time.ToShortDateString()).AddHours(8).AddMinutes(1);
            strLate = "";
            if (time < checkAM)
            {
                if (time > checkAMLate)
                {
                    strLate = "迟到";
                }
                return time;
            }
            return Convert.ToDateTime(time.ToShortDateString());
        }
        private static DateTime TimeCheckAm(DateTime time)
        {
            DateTime checkAM = Convert.ToDateTime(time.ToShortDateString()).AddHours(10).AddMinutes(1);
            if (time < checkAM)
            {
                return time;
            }
            return Convert.ToDateTime(time.ToShortDateString());
        }


        /// <summary>
        /// 方法：判断下午签退时间
        /// </summary>
        /// <param name="time">打卡时间</param>
        /// <returns></returns>
        private static DateTime TimeCheckPm(DateTime time, out string strLE)
        {
            DateTime checkPM = Convert.ToDateTime(time.ToShortDateString()).AddHours(14).AddMinutes(1);
            DateTime checkPMLE = Convert.ToDateTime(time.ToShortDateString()).AddHours(17);
            strLE = "";
            if (time > checkPM)
            {
                if (time < checkPMLE)
                {
                    strLE = "早退";
                }
                return time;
            }
            return Convert.ToDateTime(time.ToShortDateString());
        }

        //车间判断下班时间
        private static DateTime TimeCheckPm(DateTime time)
        {
            DateTime checkPM = Convert.ToDateTime(time.ToShortDateString()).AddHours(10).AddMinutes(2);

            if (time > checkPM)
            {
                return time;
            }
            return Convert.ToDateTime(time.ToShortDateString());
        }

        /// <summary>
        /// 方法：判断中午签到时间:11:59-13:30之间
        /// </summary>
        /// <param name="times">集合中所有的打卡时间</param>
        /// <returns></returns>
        private static DateTime TimeCheckMm(List<DateTime> times, out string strMmLate)
        {

            DateTime checkMM1 = Convert.ToDateTime(times[0].ToShortDateString()).AddHours(11).AddMinutes(59);
            DateTime checkMM2 = Convert.ToDateTime(times[0].ToShortDateString()).AddHours(13).AddMinutes(30);
            DateTime checkMMLate = Convert.ToDateTime(times[0].ToShortDateString()).AddHours(13).AddMinutes(1);
            strMmLate = "";
            List<DateTime> timeNoons = new List<DateTime>();
            foreach (DateTime item in times)
            {
                if (item >= checkMM1 && item <= checkMM2)
                {
                    timeNoons.Add(item);
                }
            }
            if (timeNoons.Count > 0)
            {
                foreach (var item in timeNoons)
                {
                    if (timeNoons.Min() > checkMMLate)
                    {
                        strMmLate = "迟到";
                    }
                }
                return timeNoons.Min();
            }

            return Convert.ToDateTime(times[0].ToShortDateString());



        }




    }
}
