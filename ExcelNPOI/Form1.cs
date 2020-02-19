using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using AutoUpdaterDotNET;

namespace ExcelNPOI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public List<int> Sum = new List<int>();
        public delegate int addProgress(int i);

        /// <summary>
        /// 设置根据开始日期设置结束日期
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DtBegin_ValueChanged(object sender, EventArgs e)
        {

            //根据起始日期来判断结束日期。结束日期为次月28日。
            this.dtEnd.Value = this.dtBegin.Value.Month.Equals(1) ? this.dtBegin.Value.AddMonths(1) : this.dtBegin.Value.AddMonths(1).AddDays(-1);
            
          
        }

        /// <summary>
        /// 窗体加载时的事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.lblVer.Text += version;

            //自动升级。
            //AutoUpdater.Start("ftp://192.168.0.158/versionUpdate.xml", new NetworkCredential("ftpUser1", "123"));
            AutoUpdater.Start("http://192.168.0.158:8055/update/AutoUpdater.xml");
        }

        /// <summary>
        /// 获得复选框的文本值。
        /// </summary>
        /// <returns></returns>
        public List<string> GetCBText()
        {
            List<string> cbList = new List<string>();
            foreach (CheckBox cb in this.groupBoxChoose.Controls)
            {
                if (cb.Checked)
                {
                    cbList.Add(cb.Text);
                }
            }

            return cbList;
        }

        /// <summary>
        /// 导出Excel操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnImport_Click(object sender, EventArgs e)
        {
            string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); //获取当前系统的桌面路径
            string wbName =strDesktopPath+"\\"+ this.dtEnd.Value.ToString("D") + "考勤" + "_" ;  //2019年7月28日考勤_xxxx车间

            Stopwatch swReadAccess = new Stopwatch();
            Stopwatch swWriteExcel = new Stopwatch();
            swReadAccess.Start();
            List<string> depList = GetCBText();
            if (depList.Count == 0)
            {

                MessageBox.Show("未选中任何部门、车间");
                return;

            }

            if (this.dtBegin.Value.Day ==this.dtEnd.Value.Day)
            {
                MessageBox.Show("日期天数不能相等");
                return;
            }

            if (this.dtEnd.Value.Subtract(this.dtBegin.Value).Days>31|| this.dtEnd.Value.Subtract(this.dtBegin.Value).Days<0)
            {
                MessageBox.Show("日期间隔不能大于31天，或者结束日期不能小于开始日期");
                return;
            } 

            DateTime dBegin =Convert.ToDateTime( this.dtBegin.Value.ToString("d"));
            DateTime dEnd = Convert.ToDateTime(this.dtEnd.Value.ToString("d")).AddDays(1);

            //判断是否真的要导出数据，时间和覆盖数据的说明。
            if (MessageBox.Show("是否导出选中部门的" +this.dtBegin.Value.ToString("D")+"到"+this.dtEnd.Value.ToString("D")+"的考勤数据。" +
                "导出文件将保存在桌面，并覆盖同名文件。导出时间比较长，请耐心等待一下。","重要提示",MessageBoxButtons.OKCancel)==DialogResult.Cancel)
            {
                return;
            }

            TblAttListDal dal = new TblAttListDal();

            if (depList[0]!="行政部门")
            {
                try
                {
                    //获取时间段内所有人的数据。
                    List<TblAttSource> attSource = dal.GetAttsouceQuery(dBegin, dEnd);
                    swReadAccess.Stop();

                    if (attSource == null || attSource.Count == 0)
                    {
                        MessageBox.Show("没有符合条件的考勤数据。");
                        return;
                    }

                    swWriteExcel.Start();
                    foreach (string depName in depList)
                    {
                        List<TblAttSource> listSelectSource = attSource.Where(d => d.Department == depName).ToList();

                        FuncExcel.CreateBookForCJ(this.dtBegin.Value, this.dtEnd.Value, wbName + depName + ".xlsx", listSelectSource);

                    }
                    swWriteExcel.Stop();

                    MessageBox.Show("读数据库的时间为：" + swReadAccess.Elapsed.ToString() + " ||写入Excel的时间为：" + swWriteExcel.Elapsed.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                //获取数据。并筛选非车间、前处理的数据。
                List<TblAttSource> listAttSource = dal.GetAttsouceQuery(dBegin, dEnd).Where(x=>!x.Department.Contains("车间")).Where(y=>!y.Department.Contains("前处理")).ToList();

                if (listAttSource == null || listAttSource.Count == 0)
                {
                    MessageBox.Show("没有符合条件的考勤数据。");
                    return;
                }

                swWriteExcel.Start();
                FuncExcel.CreateBookForManage(this.dtBegin.Value, this.dtEnd.Value, wbName + "行政考勤" + ".xlsx", listAttSource,"depart");

                swWriteExcel.Stop();
                swReadAccess.Stop();

                MessageBox.Show("读数据库的时间为：" + swReadAccess.Elapsed.ToString() + " ||写入Excel的时间为：" + swWriteExcel.Elapsed.ToString());

            }



            //打开保存文件对话框
            //SaveFileDialog sfd = new SaveFileDialog
            //{
            //    Title = "请选择要导出的文件路径",
            //    InitialDirectory = strDesktopPath,  //默认打开的文件夹位置。
            //    //Filter = "Excel文件|*.xls;*.xlsx|all|*.*",
            //    FileName = wbName
            //};
            //sfd.ShowDialog();
            //wbName = sfd.FileName;

            //if (wbName!="")
            //{
            //    try
            //    {
            //        FuncExcel.CreateBook(this.dtBegin.Value, this.dtEnd.Value, wbName);
            //    }
            //    catch (Exception ex)
            //    {

            //        MessageBox.Show(ex.Message);
            //        return;
            //    }
            //}

            //




        }

        /// <summary>
        /// 行政部门。选中这个checkbox，则将其他的CheckBox调整成非选中状态且不可用。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked)
            {
                foreach (CheckBox cb in this.groupBoxChoose.Controls)
                {
                    if (cb.Name!= "checkBox8")
                    {
                        cb.Checked = false;
                        cb.Enabled = false;
                    }

                }
            }
            else
            {
                foreach (CheckBox cb in this.groupBoxChoose.Controls)
                {
                    cb.Enabled = true;
                }
            }
        }

        /// <summary>
        /// 掉出其他功能窗体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFunc_Click(object sender, EventArgs e)
        {
            FormFunc formFunc = new FormFunc(this.dtBegin.Value,this.dtEnd.Value);
            formFunc.ShowDialog();
           

        }

        /// <summary>
        /// 统计结束日期日的当前在岗人数。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCountEmp_Click(object sender, EventArgs e)
        {
            TblAttListDal dal = new TblAttListDal();
            DateTime dEnd = Convert.ToDateTime(this.dtEnd.Value.ToShortDateString());
            MessageBox.Show("在统计人数之前，你应该确保已经下载了【结束日期】的所有考勤数据。", "提示");
            MessageBox.Show(this.dtEnd.Value.ToLongDateString()+"共计 "+dal.CountDateEmps(dEnd).ToString()+" 人打卡出勤。");

        }
    }
}
