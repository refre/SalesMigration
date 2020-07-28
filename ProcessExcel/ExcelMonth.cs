using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ista.Migration.Excel
{
    class ExcelMonth
    {
        public static string[] PageName
        {
            get { return pageName; }
            //set { pageName = value; }
        }   
        //private static string[] pageName = new string[] { "JAN", "FEV", "MAR", "AVR", "MAI", "JUN", "JUL", "AOU", "SEP", "OCT", "NOV", "DEC" };
        private static string[] pageName = new string[] { "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC" };
    }
}
