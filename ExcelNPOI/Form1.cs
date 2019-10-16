using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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


        private void Button1_Click(object sender, EventArgs e)
        {
            Thread progressthread = new Thread(new ParameterizedThreadStart(thread));

            progressthread.Start();
        }

        public void thread(object length)
        {
            progressForm progress = new progressForm();

            progress.Show();

            for (int i = 0; i < 100; i++)
            {
                progress.Addprogess();
                Thread.Sleep(50);
            }
            progress.Close();

        }

        public int add(int i)
        {
            Sum.Add(i);

            return Sum.Count();
        }


        private void DtBegin_ValueChanged(object sender, EventArgs e)
        {

            //根据起始日期来判断结束日期。结束日期为次月28日。
            this.dtEnd.Value = this.dtBegin.Value.Month.Equals(1) ? this.dtBegin.Value.AddMonths(1) : this.dtBegin.Value.AddMonths(1).AddDays(-1);
            
          
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //TblAttListDal dal = new TblAttListDal();
            //this.dataGridView1.DataSource = dal.GetAttListUserInfo();
        }

        //获得复选框的文本值
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
                    //读取数据库中规范好的考勤数据


                    List<TblAttList> attLists = dal.ImportForExcelList(dBegin, dEnd);
                    List<TblAttSource> attSource = dal.GetAttSource();
                    swReadAccess.Stop();

                    if (attLists == null || attLists.Count == 0)
                    {
                        MessageBox.Show("没有符合条件的考勤数据。");
                        return;
                    }

                    swWriteExcel.Start();
                    foreach (string depName in depList)
                    {

                        List<TblAttList> listSelect = attLists.Where(d => d.Dep == depName).ToList();
                        List<TblAttSource> listSelectSource = attSource.Where(d => d.Department == depName).ToList();

                        FuncExcel.CreateBook(this.dtBegin.Value, this.dtEnd.Value, wbName + depName + ".xlsx", listSelect, listSelectSource);

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
                FuncExcel.CreateBookForManage(this.dtBegin.Value, this.dtEnd.Value, wbName + "行政考勤" + ".xlsx", listAttSource);

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
        /// 方法：判断早上签到时间
        /// </summary>
        /// <param name="time">打卡时间</param>
        /// <returns></returns>
        private DateTime TimeCheckAm(DateTime time)
        {
            DateTime checkAM = Convert.ToDateTime(time.ToShortDateString()).AddHours(11).AddMinutes(1);
            if (time<checkAM)
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
        private DateTime TimeCheckPm(DateTime time)
        {
            DateTime checkPM = Convert.ToDateTime(time.ToShortDateString()).AddHours(14).AddMinutes(1);
            if (time > checkPM)
            {
                return time;
            }
            return Convert.ToDateTime(time.ToShortDateString());
        }

        /// <summary>
        /// 方法：判断中午签到时间
        /// </summary>
        /// <param name="times">集合中所有的打卡时间</param>
        /// <returns></returns>
        private DateTime TimeCheckMm(List<DateTime> times)
        {
            
            DateTime checkMM1 = Convert.ToDateTime(times[0].ToShortDateString()).AddHours(11).AddMinutes(59);
            DateTime checkMM2 = Convert.ToDateTime(times[0].ToShortDateString()).AddHours(13).AddMinutes(30);
            List<DateTime> timeNoons=new List<DateTime>();
            foreach (DateTime item in times)
            {
                if (item>=checkMM1 && item<=checkMM2)
                {
                    timeNoons.Add(item);
                }
            }
            if (timeNoons.Count>0)
            {
                return timeNoons.Min();
            }

            return Convert.ToDateTime(times[0].ToShortDateString());



        }



        private void Button2_Click(object sender, EventArgs e)
        {
            progressForm progress = new progressForm();
            progress.Show();
            for (int i = 0; i < 100; i++)
            {

                progress.Addprogess();
                Thread.Sleep(50);

            }
            progress.Close();
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

        private void button1_Click_1(object sender, EventArgs e)
        {
            string ss = "08:59:00";
            if (int.Parse(ss.Substring(0,2))>=22)
            {
                MessageBox.Show("OKI");
            }
        }
    }
}
