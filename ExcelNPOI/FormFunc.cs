using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelNPOI
{
    public partial class FormFunc : Form
    {
        private readonly DateTime _dBegin;
        private readonly DateTime _dEnd;
        public FormFunc(DateTime dBegin,DateTime dEnd)
        {
            InitializeComponent();
            this._dBegin = dBegin;
            this._dEnd = dEnd;

        }

        private void BtnMoveLastYear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("此操作将会把当前日期的上一年度数据单独归档。", "重要提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {

                if (_dBegin > DateTime.Now)
                {
                    MessageBox.Show("【结束日期】的值不能大于当前的系统日期。");
                    return;
                }


                MessageBox.Show("施工中.....");

            }
        }

        /// <summary>
        /// 导出当前周期内所有人员的出勤数据。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnImporAllDep_Click(object sender, EventArgs e)
        {
            string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); //获取当前系统的桌面路径
            string wbName = strDesktopPath + "\\" + _dEnd.ToString("D") + "考勤" + "_";  //2019年7月28日考勤_xxxx车间

            Stopwatch swReadAccess = new Stopwatch();
            Stopwatch swWriteExcel = new Stopwatch();
            swReadAccess.Start();
            

            if (_dBegin.Day == _dEnd.Day)
            {
                MessageBox.Show("日期天数不能相等");
                return;
            }

            if (_dEnd.Subtract(_dBegin).Days > 31 || _dEnd.Subtract(_dBegin).Days < 0)
            {
                MessageBox.Show("日期间隔不能大于31天，或者结束日期不能小于开始日期");
                return;
            }

            DateTime dBegin = Convert.ToDateTime(_dBegin.ToString("d"));
            DateTime dEnd = Convert.ToDateTime(_dEnd.ToString("d")).AddDays(1);

            //判断是否真的要导出数据，时间和覆盖数据的说明。
            if (MessageBox.Show("是否导出所有人员的" + _dBegin.ToString("D") + "到" + _dEnd.ToString("D") + "的考勤数据。" +
                "导出文件将保存在桌面，并覆盖同名文件。导出时间比较长，请耐心等待一下。", "重要提示", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
            {
                return;
            }

            TblAttListDal dal = new TblAttListDal();
            //获取数据
            List<TblAttSource> listAttSource = dal.GetAttsouceQuery(dBegin, dEnd).ToList();

            if (listAttSource == null || listAttSource.Count == 0)
            {
                MessageBox.Show("没有符合条件的考勤数据。");
                return;
            }

            swWriteExcel.Start();
            FuncExcel.CreateBookForManage(_dBegin, _dEnd, wbName + "全部人员" + ".xlsx", listAttSource,"");

            swWriteExcel.Stop();
            swReadAccess.Stop();

            MessageBox.Show("读数据库的时间为：" + swReadAccess.Elapsed.ToString() + " ||写入Excel的时间为：" + swWriteExcel.Elapsed.ToString());



        }
    }
}
