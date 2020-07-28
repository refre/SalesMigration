using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Ista.Migration.Excel
{
    /// <summary>
    /// This class is created in order to work with the ecxel file named "Offres".
    /// It can be found on the server in the Sales departement.
    /// http://stackoverflow.com/questions/1164698/using-excel-oledb-to-get-sheet-names-in-sheet-order
    /// </summary>
    public class OffresExcel: Excel12Manager
    {
        /// <summary>
        /// Initialzes a new OfferExcel class.
        /// </summary>
        /// <param name="path">Path of the file "Offres".</param>
        public OffresExcel(string path) : base(path)
        {
            
        }

        /// <summary>
        /// Read a single worksheet and transform it in datatable.
        /// </summary>
        /// <param name="workSheetName">Current Worksheet name.</param>
        /// <returns>Datatable containing values stored in Excel.</returns>
        public override DataTable ReadWorkSheet(string workSheetName)
        {
            string currentWorkSheet = workSheetName + "$";

            string strSQL = "SELECT * FROM [" + currentWorkSheet + "]";
            DataTable dTable = new DataTable();

            using (OleDbConnection excelConnection = new OleDbConnection(base._sbConnection.ToString()))
            {
                excelConnection.Open();
                OleDbCommand dbCommand = new OleDbCommand(strSQL, excelConnection);
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(dbCommand);
                // create data table            
                dataAdapter.Fill(dTable);
                excelConnection.Close();
            }
            return dTable;
        }

        /// <summary>
        /// Gets a list of datatable containing all the pages in Excel.
        /// </summary>
        /// <returns>Lsit of datatable.</returns>
        public List<DataTable> GetCompleteExcelInList()
        {
            Stopwatch timeElapsed = new Stopwatch();
            timeElapsed.Start();
            List<DataTable> ReturnedDatatable = new List<DataTable>();
            Parallel.ForEach(ExcelMonth.PageName, currentTable => 
                ReturnedDatatable.Add( ReadWorkSheet(currentTable)));

            timeElapsed.Stop();

            TimeSpan timeElapsedValue = timeElapsed.Elapsed;

            return ReturnedDatatable;
        }
    }
}
