using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ista.Migration.Excel;
using System.Data;
using System.Collections.ObjectModel;

namespace ProcessExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            OffresExcel excelDoc = new OffresExcel(@"C:\test\Copie - BX - Offres-cmd-2012.xlsx");
            DataTable myTable = excelDoc.ReadWorkSheet("JAN");

            int counter = 0;
            for (int i = 0; i < myTable.Rows.Count; i++)
            {
                if(!string.IsNullOrEmpty(myTable.Rows[i][0].ToString()))
                {
                    counter++;
                }
                else
            	{
                    break;
	            }
            }

            ExcelExportHelper myWriteExcel = new ExcelExportHelper(@"C:\test\BX - Offres-cmd-2012 - Copie.xlsx");
            myWriteExcel.ChangeSheet("JAN");
            myWriteExcel.SetCell(79, 3, "ValueAdded");
            myWriteExcel.Save(@"C:\test\BX - Offres-cmd-2012 - Copie.xlsx");
            myWriteExcel.Dispose();

            Console.ReadLine();

        }
    }
}
