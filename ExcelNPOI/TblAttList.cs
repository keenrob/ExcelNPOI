using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNPOI
{
    public partial class TblAttList
    {
 
        public int UserID { get; set; }
        public string UName { get; set; }
        public string Dep { get; set; }
        public DateTime AttDate { get; set; }
        public string Time1 { get; set; }
        public  string Time2 { get; set; }
        public int TimeDiff { get; set; }
        public string CheckAtt { get; set; }
        public double AttHours { get; set; }



    }
}
