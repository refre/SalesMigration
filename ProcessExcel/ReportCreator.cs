using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;


namespace Ista.Migration.Excel
{
    /// <summary>
    /// This class creates a report for the buildings.
    /// </summary>
    public class ReportCreator
    {
        /// <summary>
        /// Private variable: Hastable of the excel process
        /// </summary>
        private Hashtable _myHashtable;
        private const string _newbxFilePath = @"C:\Users\kamnanjo\Downloads\Edourad\2013-8\BX Michael - Offres-cmd-2013.xlsx";
        private const string _newanFilePath = @"C:\Users\kamnanjo\Downloads\Edourad\2013-8\AN Karel - Offres-cmd-2013.xlsx";
        private const string _newvvFilePath = @"C:\Users\kamnanjo\Downloads\Edourad\2013-8\VV Benoit - offres-cmd-2013.xlsx";

        /// <summary>
        /// Initialize a new ReportCreator
        /// </summary>
        /// <param name="migrationList">Building list.</param>
        /// <param name="bxFilePath">Path of the file for brussels.</param>
        /// <param name="vvFilePath">Path of the file for Verviers.</param>
        /// <param name="anFilePath">Path of the file for Antwerp.</param>
        public ReportCreator(List<MigrationElement> migrationList, string bxFilePath, string anFilePath, string vvFilePath)
        {
            //1) We must first segregate any city with the zipcode
            string[] antw = new string[] { "2", "8", "90", "91", "92", "97", "98", "99", "35", "36", "37", "39" };
            string[] brux = new string[] { "30", "31", "32", "33", "34", "38", "1", "7", "60", "61", "62", "64","65", "93", "94", "95", "96" };
            string[] Verv = new string[] { "4", "5", "66", "67", "68","69" };
            
            
            List<MigrationElement> bxMigration = new List<MigrationElement>();
            List<MigrationElement> anMigration = new List<MigrationElement>();
            List<MigrationElement> vvMigration = new List<MigrationElement>();

            CheckExcellProcesses();

            foreach (var item in migrationList)
            {
                if (item.CodePostalImmeuble.StartsWithAny(brux))
                {
                    bxMigration.Add(item);                    
                }
            }
            foreach (var item in migrationList)
            {
                if (item.CodePostalImmeuble.StartsWithAny(antw))
                {
                    anMigration.Add(item);                    
                } 
            }
            foreach (var item in migrationList)
            {
                if (item.CodePostalImmeuble.StartsWithAny(Verv))
                {
                    vvMigration.Add(item);   
                }
            }

            string currentMonth = ExcelMonth.PageName[DateTime.Now.Month - 1];
            
            if (bxMigration.Count > 0)
            {
                ExcelExportHelper bxExcel = new ExcelExportHelper(bxFilePath);
                //CityProcess(bxExcel, bxFilePath, currentMonth, bxMigration,"MSM");
                CityProcess(bxExcel, bxFilePath, currentMonth, bxMigration, "MSM");
            }
            if (anMigration.Count > 0)
            {
                ExcelExportHelper anExcel = new ExcelExportHelper(anFilePath);
                //CityProcess(anExcel, anFilePath, currentMonth, anMigration, "KMA");
                CityProcess(anExcel, anFilePath, currentMonth, anMigration, "KMA");
            }
            if (vvMigration.Count > 0)
            {
                ExcelExportHelper vvExcel = new ExcelExportHelper(vvFilePath);
                //CityProcess(vvExcel, vvFilePath, currentMonth, vvMigration, "BGI");
                CityProcess(vvExcel, vvFilePath, currentMonth, vvMigration, "BGI");
            }
            KillExcel();
        }
        /// <summary>
        /// This method process the value into the cell of the Excel file.
        /// </summary>
        /// <param name="excelExport">New Excel File.</param>
        /// <param name="filePath">Path of the file</param>
        /// <param name="currentMonth">Current Month.</param>
        /// <param name="migration">Migration data.</param>
        private void CityProcess(ExcelExportHelper excelExport, string filePath,string currentMonth, List<MigrationElement> migration,string sales)
        {
            int excelRowCount = NomberOfRowUsed(filePath, currentMonth);
            excelExport.ChangeSheet(currentMonth);

           foreach (var item in migration)
            {
                CoreBuiness(excelExport, item, excelRowCount,sales);
                excelRowCount++;
            }

            string directory = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            excelExport.Save(directory + @"\tempbx.xlsx");
            excelExport.Close();
            excelExport.Dispose();

            BackupOldExcelFile(filePath);

            File.Copy(directory + @"\tempbx.xlsx", directory +@"\" +fileName + ".xlsx", true);
            File.Delete(directory + @"\tempbx.xlsx");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        private void BackupOldExcelFile(string filePath)
        {
            string directory = Path.GetDirectoryName(filePath);
            string newFileName = Path.GetFileNameWithoutExtension(filePath)+DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")+".xlsx";
            string newDirectory = directory + @"\backup";
            string newFileInDirectory = newDirectory + @"\" + newFileName;

            if (!Directory.Exists(newDirectory))
            {
                Directory.CreateDirectory(newDirectory);
            }
            try
            {
                File.Copy(filePath,  newFileInDirectory);
            }
            catch (UnauthorizedAccessException uaex)
            {
                throw new UnauthorizedAccessException("The user doesnt have enough priviledge to move the file", uaex);
            }
            catch (Exception ex)
            {
                throw new Exception("Error", ex);
            }
        }

        /// <summary>
        /// Find the first row where the data could be written.
        /// </summary>
        /// <param name="path">Path of the file where the numbers of rows are contained.</param>
        /// <param name="month">Name of the month where the offer will be written.</param>
        /// <returns>First empty row.</returns>
        private int NomberOfRowUsed(string path, string month)
        {
            OffresExcel excelDoc = new OffresExcel(path);
            DataTable myTable = excelDoc.ReadWorkSheet(month);

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
            myTable.Dispose();

            // I add the +2 because 1) A datatable is Zero base indexeb but Excel is 1 bas indexed (+1)
            // The Excel file has a first row containing Data (+1) 
            // Total is +2 
            return counter+2;
        }
        /// <summary>
        /// Method the process the value to Excel
        /// </summary>
        /// <param name="currentExcel">Excel file instance.</param>
        /// <param name="currentElement">Current data element.</param>
        /// <param name="rowCount">The row where the data will be placed.</param>
        private void CoreBuiness(ExcelExportHelper currentExcel,MigrationElement currentElement,int rowCount,string saleseName)
        {
           

            //int tot = currentElement.Element.TotEau + currentElement.Element.TotInteg + currentElement.Element.TotNchauf;

            currentExcel.InsertRows(rowCount);
            
            if (currentElement.Element.TotalDevice < 51)
            {
                for (int i = 1; i < 40; i++)
                {
                    currentExcel.SetLineColor(rowCount, i);
                }
            }
            else
            {
                for (int i = 1; i < 40; i++)
                {
                    currentExcel.SetSecondLineColor(rowCount, i);
                }
            }

            currentExcel.SetCell(rowCount, 1, saleseName);
            currentExcel.SetCell(rowCount, 2, currentElement.Site);
            currentExcel.SetCell(rowCount, 4, currentElement.NumeroImmeuble);
            currentExcel.SetCell(rowCount, 6, currentElement.Element.TotalAdressImeuble);
            currentExcel.SetCell(rowCount, 7, currentElement.CodePostalImmeuble);
            currentExcel.SetCell(rowCount, 8, currentElement.LocaliteImmeuble);
            currentExcel.SetCell(rowCount, 9, currentElement.Element.CodeNameGerant);
            currentExcel.SetCell(rowCount, 10, currentElement.ChauffageNombreNRad);
            currentExcel.SetCell(rowCount, 11, "DR3");
            currentExcel.SetCell(rowCount, 12, currentElement.Element.Dop3RadioChaufLocatValue);
            currentExcel.SetCell(rowCount, 13, currentElement.Element.Dop3RadioChaufVenteValue);
            currentExcel.SetCell(rowCount, 14, currentElement.Element.Dop3RadioChaufRelevValue);

            string TypeText = "";
            if (currentElement.Element.TotEau > 0)
            {
                if (currentElement.Element.TotNEauch > 0 & currentElement.Element.TotNEauFr == 0)
                {
                    TypeText = "EC";
                }
                else if (currentElement.Element.TotNEauch == 0 & currentElement.Element.TotNEauFr > 0)
                {
                    TypeText = "EF";
                }
                else
                {
                    TypeText = "EF EC";
                }
            }

            currentExcel.SetCell(rowCount, 15, currentElement.Element.TotEau);
            currentExcel.SetCell(rowCount, 16, TypeText);
            currentExcel.SetCell(rowCount, 17, currentElement.Element.DomaquaTotalLocatiValue);
            currentExcel.SetCell(rowCount, 18, currentElement.Element.DomaquaTotalVenteValue);
            currentExcel.SetCell(rowCount, 19, currentElement.Element.DomaquaTotalReleveValue);

            currentExcel.SetCell(rowCount, 20, currentElement.Element.TotInteg);
            currentExcel.SetCell(rowCount, 21, "RADIO");
            currentExcel.SetCell(rowCount, 22, currentElement.Element.SensonicIn1_2LocatValue);
            currentExcel.SetCell(rowCount, 23, currentElement.Element.SensonicIn1_2VenteValue);
            currentExcel.SetCell(rowCount, 24, currentElement.Element.SensonicIn1_2RelevValue);

            currentExcel.SetCell(rowCount, 28, currentElement.DocumentName);
            currentExcel.SetCell(rowCount, 29, DateTime.Now.ToString("dd/MM/yyyy"));

        }
        /// <summary>
        /// This method checks whether there are already some excel process running.
        /// </summary>
        private void CheckExcellProcesses()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            _myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process ExcelProcess in AllProcesses)
            {
                _myHashtable.Add(ExcelProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }
        /// <summary>
        /// This method kills excel process.
        /// </summary>
        private void KillExcel()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            // check to kill the right process
            foreach (Process ExcelProcess in AllProcesses)
            {
                if (_myHashtable.ContainsKey(ExcelProcess.Id) == false)
                    ExcelProcess.Kill();
            }
            AllProcesses = null;
        }
    }   
}
