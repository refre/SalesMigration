using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Media;
using System.Runtime.InteropServices;

namespace Ista.Migration.Excel
{
    public class ExcelExportHelper: IDisposable
    {
        private dynamic _objXLApp;
        private dynamic _objXLBook;
        private dynamic _objXLSheet;
        private string _strHeader;
        private System.Globalization.CultureInfo _originalCulture;
        private enum excel : short
        {
            XLExpression = 2,
            XLNONE = -4142,
            XLCONTINUOUS = 1,
            XLTHIN = 2,
            XLEdgeLeft = 7,
            XLEdgeRight = 10,
            XLEdgeTop = 8,
            XLEdgeBottom = 9,
            XLInsideVertical = 12,
            XLInsideHorizontal = 11,
            XLDiagonalUp = 6,
            XLDiagonalDown = 5,
            XLAutomatic = -4105
        }

        #region Properties

        /// <summary>
        ///READ ONLY PROPERTY
        /// The XL application used by this object.
        /// </summary>
        /// <value>The XL application used by this object</value>
        /// <returns>The XL application used by this object</returns>
        /// <remarks></remarks>
        public dynamic Application
        {
            // get instance of the Application object
            get { return _objXLApp; }
        }

        /// <summary>
        /// The header of this object (used by the worksheet)
        /// </summary>
        /// <value>The header of this object</value>
        /// <returns>The header of this object</returns>
        /// <remarks></remarks>
        public string Header
        {
            get { return _strHeader; }
            set { _strHeader = value; }
        }

        /// <summary>
        /// The cultureInfo in use at the instantiation of this object.
        /// </summary>
        /// <value>The cultureInfo in use at the instantiation of this object</value>
        /// <returns>The cultureInfo in use at the instantiation of this object</returns>
        /// <remarks></remarks>
        public System.Globalization.CultureInfo OriginalCulture
        {
            get { return _originalCulture; }
        }

        #endregion

        #region Constructor(s)

        /// <summary>
        /// Creates a NEW workbook in memory.  (This call will change the culture
        /// prior to instantiating the XL objects !)
        /// </summary>
        /// <remarks>This call will change the cultureprior to instantiating the XL objects!</remarks>
        public ExcelExportHelper()
        {
            _originalCulture = ExcelExportHelper.GetCurrentCulture();
            ExcelExportHelper.SetExcelCulture();

            // create new instance of the Excell application and create one worksheet
            _objXLApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            _objXLBook = _objXLApp.workbooks.add();
            _objXLSheet = _objXLBook.worksheets(1);
        }

        /// <summary>
        /// Opens the given file in memory.  
        /// Note: 
        /// 1) This call will change the culture prior to instantiating the XL objects !
        /// 2) if it doesn't exist yet, an exception will be thrown.
        /// </summary>
        /// <param name="strFileName">The file to open.  (if it doesn't exist yet, an
        /// exception will be thrown.)</param>
        /// <remarks></remarks>
        public ExcelExportHelper(string strFileName)
        {
            _originalCulture = ExcelExportHelper.GetCurrentCulture();
            ExcelExportHelper.SetExcelCulture();
            // create new instance of the XLS application and open existing worksheet
            _objXLApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            _objXLBook = _objXLApp.workbooks.open(strFileName);
            _objXLSheet = _objXLBook.worksheets(1);
        }

        #endregion

        #region Methods

        #region --Public

        #region ----Operations on worksheets

        /// <summary>
        /// Gets the current sheet.  You may also use this so get to another worksheet
        /// by combining the changeSheet method followed by the getActiveSheet method.
        /// </summary>
        /// <returns>The active worksheet object (not the name but well the object itself)</returns>
        /// <remarks></remarks>
        public object GetActiveSheet()
        {
            object SheetToReturn = _objXLSheet;
            return SheetToReturn;
        }

        /// <summary>
        /// Method to insert a new row in active sheet
        /// </summary>
        /// <param name="intRowNum"></param>
        public void InsertRows(int intRowNum)
        {
            //dynamic rng = _objXLSheet.Range("A" + intRowNum.ToString());
            //rng = rng.EntireRow;
            //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
        }

        /// <summary>
        /// Returns an arraylist holding the names of all worksheets available in the 
        /// workbook.
        /// </summary>
        /// <returns>arraylist holding the names of all worksheets available in the 
        /// workbook</returns>
        /// <remarks></remarks>
        public ArrayList GetSheetNames()
        {
            ArrayList sheetList = new ArrayList();
            foreach (var sheet in _objXLBook.worksheets)
            {
                sheetList.Add(sheet.Name);
            }
            return sheetList;
        }

        /// <summary>
        /// Modifies the currently selected worksheet.
        /// </summary>
        /// <param name="strNameOfSheet">The name of the worksheet that you wish to select</param>
        /// <returns>True if the worksheet was found.  False otherwise</returns>
        /// <remarks></remarks>
        public bool ChangeSheet(string strNameOfSheet)
        {
            foreach (var sheet in _objXLBook.worksheets)
            {
                if (sheet.Name /*.tolower*/ == strNameOfSheet/*.ToLower()*/ | sheet.CodeName/*.toLower*/ == strNameOfSheet/*.ToLower()*/)
                {
                    _objXLSheet = sheet;
                    return true;
                }
            }

            return false;
        }

        /// <summary>Adds a worksheet to the current workbook and sets it as the current worksheet
        /// </summary>
        /// <param name="worksheetName">the name of the worksheet</param>
        public void AddWorksheet(string worksheetName)
        {
            _objXLSheet = _objXLBook.Worksheets.Add();
            _objXLSheet.Name = worksheetName;
        }

        #endregion

        #region ----Operations on cells

        /// <summary>
        /// Returns the cell dynamic object (= an Excel Range object).  
        /// </summary>
        /// <param name="intRowNum">The row of the cell (starts at 1)</param>
        /// <param name="intColNum">The column of the cell (starts at 1)</param>
        /// <returns></returns>
        public dynamic GetCell(int intRowNum, int intColNum)
        {
            return _objXLSheet.Range(GetColumnName(intColNum) + intRowNum.ToString());
        }

        /// <summary>
        /// Returns the cell dynamic object (= an Excel Range object).  
        /// </summary>
        /// <param name="rangeName">The name of the cell.  It can be in the format A1, B6, ... or a named range such as
        /// computedTotal, NumberOfMonth, ...</param>
        /// <returns></returns>
        public dynamic GetCell(string rangeName)
        {
            return _objXLSheet.range(rangeName).value;
        }

        /// <summary>
        /// Sets the value of a cell.  Note that the value is declared as object so that you can pass
        /// numeric or date values without conerting them to string!
        /// </summary>
        /// <param name="intRowNum">The row of the cell (starts at 1)</param>
        /// <param name="intColNum">The column of the cell (starts at 1)</param>
        /// <param name="value">The value you wish to write in the cell</param>
        public void SetCell(int intRowNum, int intColNum, Object value)
        {
            dynamic cell = GetCell(intRowNum, intColNum);
            if (value == null)
            {
                cell.Value = string.Empty;
            }
            else
            {
                try
                {
                    cell.Value = value;
                }
                catch (Exception)
                {
                    cell.Value = value.ToString();
                }
            }
        }

        /// <summary>
        /// Sets the value of a cell.  Note that the value is declared as object so that you can pass
        /// numeric or date values without conerting them to string!
        /// </summary>
        /// <param name="rangeName">The name of the range in which you wish to write the value.  It can be in the format 
        /// A1, B6, ... or a named range such as computedTotal, NumberOfMonth, ...</param>
        /// <param name="value">The value you wish to write in the cell/range.</param>
        public void SetCell(string rangeName, object value)
        {
            dynamic cell = GetCell(rangeName);
            if (value == null)
            {
                cell.Value = string.Empty;
            }
            else
            {
                try
                {
                    cell.Value = value;
                }
                catch (Exception)
                {
                    cell.Value = value.ToString();
                }
            }
        }

        #endregion


        /// <summary>
        /// Function to do some proper clean up of objects accessed via OLE
        /// </summary>
        /// <remarks></remarks>
        void System.IDisposable.Dispose()
        {
            // reset the XLS objects
            _objXLSheet = null;
            _objXLBook = null;
            _objXLApp = null;
            //note that the application is not closed here because in some cases, you'll want to display the generated document and consequently, the application should not be closed

            ExcelExportHelper.RestoreCulture(OriginalCulture);
        }

        /// <summary>
        /// Frees loaded excel objects and restores the culture
        /// </summary>
        public void Dispose()
        {
            ((System.IDisposable)this).Dispose();
        }

        /// <summary>
        /// Closes the workbook from memory, exits the application, cleans up the objects, and
        /// restores the culture.
        /// </summary>
        /// <remarks></remarks>
        public void Close()
        {
            //Closes the worksheet , workbook and application without saving
            _objXLSheet = null;
            if ((_objXLBook != null))
            {
                _objXLBook.close(false);
            }

            _objXLApp.quit();
            Marshal.ReleaseComObject(_objXLApp);
            

            //Now, restore the culture
            ExcelExportHelper.RestoreCulture(_originalCulture);

        }

        /// <summary>
        /// sets an alternate row style on all rows of teh worksheet in memory
        /// </summary>
        /// <param name="Oddcolor">USELESS ! this method always uses the Burgundy bright color
        /// for the odd rows</param>
        /// <remarks></remarks>
        public void SetPijama(Color Oddcolor)
        {
            // create the alternate row style (Pijama) in the XLS
            // ---------------------------------------------------
            int intColor = 0;

            try
            {
                intColor = (256 * 256 * Oddcolor.B) + (256 * Oddcolor.G) + Oddcolor.R;
                {
                    _objXLSheet.activate();
                    _objXLSheet.usedRange.Rows.Select();
                    _objXLSheet.application.Selection.FormatConditions.Delete();
                    _objXLSheet.application.Selection.FormatConditions.Add(excel.XLExpression, null, "=MOD(ROW(A1);2)");
                    _objXLSheet.application.selection.FormatConditions(1).Interior.Color = intColor;

                    _objXLSheet.range("A1").Select();
                }


            }
            catch (Exception ex)
            {
                throw new Exception("Cannot set pijama in excel", ex);
            }
        }

        /// <summary>
        /// Put line in color
        /// </summary>
        /// <param name="RowNumber"></param>
        /// <param name="col"></param>
        public void SetLineColor(int RowNumber, int col)
        {
            try
            {
                int color = (256 * 256 * 127) + (256 * 132) + 235;
                {
                    _objXLSheet.Cells[RowNumber, col].Interior.Color = color; 
                }
            }
            catch (Exception ex)
            {

                throw new Exception("Cannot set colors", ex); 
            }           
        }

        /// <summary>
        /// put line in color referenced
        /// </summary>
        /// <param name="RowNumber"></param>
        /// <param name="col"></param>
        /// <param name="color"></param>
        private void SetLineColorBase(int RowNumber, int col,int color)
        {
            try
            {  
                _objXLSheet.Cells[RowNumber, col].Interior.Color = color;
            }
            catch (Exception ex)
            {

                throw new Exception("Cannot set colors", ex);
            }
        }

        /// <summary>
        /// put line in color referenced
        /// </summary>
        /// <param name="RowNumber"></param>
        /// <param name="col"></param>
        /// <param name="color"></param>
        public void SetSecondLineColor(int rowNumber, int col)
        {
            int color = (256*256*32) + (256*144) + 89;
            SetLineColorBase(rowNumber, col, color);
        }


        ///<summary> Well, the method speaks by itself.
        /// </summary>
        /// <remarks></remarks>
        public void AutoFitColumns()
        {
            // adjust size of columns to autofit
            _objXLSheet.columns.autofit();

        }

        /// <summary>
        /// Sets the LEFT header of the active worksheet
        /// The font used is Arial bold.
        /// </summary>
        /// <param name="strHeaderText">The text you want displayed</param>
        /// <remarks></remarks>
        public void SetHeader(string strHeaderText)
        {
            // set header text
            _objXLSheet.PageSetup.LeftHeader = "&\"Arial,Bold\"" + strHeaderText;

        }

        /// <summary>
        /// Sets top, bottom, left and right borders arround all cells of the active worksheet
        /// </summary>
        /// <remarks></remarks>
        public void SetBorders()
        {
            // set border of the XLS cells (range)
            {
                _objXLSheet.usedRange.borders(excel.XLDiagonalDown).lineStyle = excel.XLNONE;
                _objXLSheet.usedRange.borders(excel.XLDiagonalUp).lineStyle = excel.XLNONE;

                {
                    _objXLSheet.usedRange.borders(excel.XLEdgeLeft).LineStyle = excel.XLCONTINUOUS;
                    _objXLSheet.usedRange.borders(excel.XLEdgeLeft).weight = excel.XLTHIN;
                    _objXLSheet.usedRange.borders(excel.XLEdgeLeft).ColorIndex = excel.XLAutomatic;
                }
                {
                    _objXLSheet.usedRange.borders(excel.XLEdgeRight).LineStyle = excel.XLCONTINUOUS;
                    _objXLSheet.usedRange.borders(excel.XLEdgeRight).weight = excel.XLTHIN;
                    _objXLSheet.usedRange.borders(excel.XLEdgeRight).ColorIndex = excel.XLAutomatic;
                }
                {
                    _objXLSheet.usedRange.borders(excel.XLEdgeTop).LineStyle = excel.XLCONTINUOUS;
                    _objXLSheet.usedRange.borders(excel.XLEdgeTop).weight = excel.XLTHIN;
                    _objXLSheet.usedRange.borders(excel.XLEdgeTop).ColorIndex = excel.XLAutomatic;
                }
                {
                    _objXLSheet.usedRange.borders(excel.XLEdgeBottom).LineStyle = excel.XLCONTINUOUS;
                    _objXLSheet.usedRange.borders(excel.XLEdgeBottom).weight = excel.XLTHIN;
                    _objXLSheet.usedRange.borders(excel.XLEdgeBottom).ColorIndex = excel.XLAutomatic;
                }
                {
                    _objXLSheet.usedRange.borders(excel.XLInsideVertical).LineStyle = excel.XLCONTINUOUS;
                    _objXLSheet.usedRange.borders(excel.XLInsideVertical).weight = excel.XLTHIN;
                    _objXLSheet.usedRange.borders(excel.XLInsideVertical).ColorIndex = excel.XLAutomatic;
                }
                {
                    _objXLSheet.usedRange.borders(excel.XLInsideHorizontal).LineStyle = excel.XLCONTINUOUS;
                    _objXLSheet.usedRange.borders(excel.XLInsideHorizontal).weight = excel.XLTHIN;
                    _objXLSheet.usedRange.borders(excel.XLInsideHorizontal).ColorIndex = excel.XLAutomatic;
                }
            }
        }

        /// <summary>
        /// Saves the workbook in memory under the file name provided to the method.  It does not restore the culture !
        /// </summary>
        /// <param name="strFilename">The destination filename</param>
        /// <returns>True</returns>
        /// <remarks></remarks>
        public bool Save(string strFilename)
        {
            // save XLS file by giving the XLS file name
            _objXLBook.saveas(strFilename);
            return true;
        }

        /// <summary>
        /// Saves the workbook in memory under the file name provided to the method but it allows
        /// you to specify what the system should do if the file already exist.
        /// It does not restore the culture !
        /// </summary>
        /// <param name="strFilename">The destination file name</param>
        /// <param name="blnOverwrite">True to delete any existing file an overwrite it
        /// False otherwise</param>
        /// <returns>True if the file was successfully saved.
        /// False otherwise (note that saving is ignored if blnOverwrite = false and a file 
        /// already exists)</returns>
        /// <remarks></remarks>
        public bool Save(string strFilename, bool blnOverwrite)
        {
            // save XLS file by giving the XLS file name and indicate if overwrite of existing file (1st parameter) needs to be done
            // ----------------------------------------------------------------------------------------------------------------------
            System.IO.FileInfo fileTarget = new System.IO.FileInfo(strFilename);

            if (fileTarget.Exists)
            {
                // check if this file already exists
                if (blnOverwrite)
                {
                    fileTarget.Delete();
                }
                else
                {
                    return false;
                }
            }

            return Save(strFilename);
        }

        /// <summary>
        /// Imports a datatable into the current worksheet starting at row 1, column 1.
        /// ALL columns are imported in XL
        /// </summary>
        /// <param name="dtSource">The datatable to import into the workbook</param>
        /// <remarks></remarks>
        public void ImportData(DataTable dtSource)
        {
            //
            // Import data of a datatable to currently opened XLS file
            // -----------------------------------------------------------
            short shtColCounter = 0;
            short shtRowCounter = 0;

            //Set NumberFormat property of excel columns
            //------------------------------------------
            for (int inti = 0; inti <= dtSource.Columns.Count - 1; inti++)
            {
                DataColumn dc = dtSource.Columns[inti];

                setColumnDataType(GetColumnName(inti + 1), dc.DataType);

            }

            //Start writing the data
            //------------------------
            foreach (DataRow row in dtSource.Rows)
            {
                shtRowCounter += 1;
                shtColCounter = 0;
                foreach (DataColumn col in dtSource.Columns)
                {
                    shtColCounter += 1;
                    //YAU on 18/02/2008 : if I leave the toString, then everything is 
                    //passed as string !!

                    //cell(shtRowCounter, shtColCounter) = row.Item(col).ToString
                    SetCell(shtRowCounter, shtColCounter, row[col]);
                }
            }
        }
                
        /// <summary>
        /// Gets an arrayList holding all named ranges
        /// </summary>
        /// <returns>Arraylist of strings : each element being the name of a named range</returns>
        /// <remarks></remarks>
        public ArrayList GetRanges()
        {

            ArrayList alResult = new ArrayList();

            {
                int i = 0;
                for (i = 1; i <= _objXLBook.Names.Count; i++)
                {
                    if (!_objXLBook.Names(i).Name.ToString.Contains("!"))
                        alResult.Add(_objXLBook.Names(i).Name.ToString);
                }
            }

            return alResult;

        }

        #endregion

        #region --Private


        /// <summary>
        /// Method which sets the data type of the specified column in the current sheet
        /// </summary>
        /// <param name="strColName">The name of the column (e.g. A, B, BC, ...)</param>
        /// <param name="dataType">A type.  Only the following types are handled: Boolean, Datetime,
        /// Decimal, double, single, int32, int64, UINT16, UInt32, UInt64, String.  All other types
        /// are set to General.</param>
        /// <remarks></remarks>
        private void setColumnDataType(string strColName, Type dataType)
        {
            switch (Type.GetTypeCode(dataType))
            {
                case TypeCode.Boolean:
                    _objXLSheet.range(strColName + ":" + strColName).NumberFormat = "@";
                    break;
                case TypeCode.DateTime:
                    _objXLSheet.range(strColName + ":" + strColName).NumberFormat = "dd/mm/yyyy";
                    break;
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    _objXLSheet.range(strColName + ":" + strColName).NumberFormat = "0.00";
                    break;
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    _objXLSheet.range(strColName + ":" + strColName).NumberFormat = "0";
                    break;
                case TypeCode.String:
                    _objXLSheet.range(strColName + ":" + strColName).NumberFormat = "@";
                    break;
                default:
                    _objXLSheet.range(strColName + ":" + strColName).NumberFormat = "General";
                    break;
            }
        }


        #endregion

        #region --Static
        /// <summary>
        /// Gets the name of a column based on its column number
        /// </summary>
        /// <param name="intColNumber">The column number</param>
        /// <returns>The column name as string</returns>
        /// <example>excelExportHelper.getColumnname(4) will return D</example>
        /// <remarks></remarks>
        public static string GetColumnName(int intColNumber)
        {
            //
            // Return column name based on column number
            // example column number 26 --> return Z  (the 26 the column is named Z)
            // ---------------------------------------------------------------------
            int intIntegerPart = 0;
            int intModuloPart = 0;

            //If shtColNumber = 0 Then
            //    Return ""
            if (intColNumber <= 26)
            {
                return ((char)(64 + intColNumber)).ToString();
            }
            else
            {
                intIntegerPart = (intColNumber - 1) / 26;
                intModuloPart = 1 + ((intColNumber - 1) % 26);
                return GetColumnName(intIntegerPart) + GetColumnName(intModuloPart);
            }

        }

        /// <summary>
        /// Gets the current cultureInfo.
        /// </summary>
        /// <returns>The current cultureInfo.</returns>
        /// <remarks></remarks>
        public static System.Globalization.CultureInfo GetCurrentCulture()
        {
            //
            // return current culture of the system (user)
            // -------------------------------------------
            System.Threading.Thread thisThread = System.Threading.Thread.CurrentThread;
            System.Globalization.CultureInfo currentCulture = thisThread.CurrentCulture;

            return currentCulture;
        }

        /// <summary>
        /// Same as importData(dtSource as dataTable) except that this method can be called
        /// without instantiating an excelexportHelper object.
        /// This method automatically sizes the columns, adds borders around each cell and adds
        /// an alternating row style.
        /// </summary>
        /// <param name="dtSource">The datatable to export</param>
        /// <param name="strHeader">The left header for the worsheet</param>
        /// <remarks></remarks>
        public static void ExportTable(DataTable dtSource, string strHeader = "")
        {
            //
            // export data of a datatable to a new XLS (use of imortData sub defined above)
            // Use to export the datatable in one line (example Exporttable(dtResults)
            // -------------------------------------------------------------------------
            ExcelExportHelper xlExporter = null;

            System.Globalization.CultureInfo originalCulture = GetCurrentCulture();

            try
            {
                if (dtSource.Rows.Count == 0)
                {
                    throw new Exception("There is no row to export.  Operation aborted");
                }

                SetExcelCulture();

                xlExporter = new ExcelExportHelper();

                {
                    xlExporter.SetHeader(strHeader);
                    xlExporter.ImportData(dtSource);
                    xlExporter.AutoFitColumns();
                    xlExporter.SetBorders();

                    xlExporter.Application.visible = true;
                }

            }
            catch (Exception ex)
            {
                if (xlExporter != null)
                {
                    xlExporter.Close();
                }
                throw new Exception("Excel Export Error", ex);
            }
            finally
            {
                RestoreCulture(originalCulture);
            }
        }

        /// <summary>
        /// Changes the cuture to EN-US so that excel objects can used safely
        /// </summary>
        /// <remarks></remarks>
        public static void SetExcelCulture()
        {
            //
            // Change current culture of the XLS --> Force en-US
            // --------------------------------------------------
            System.Threading.Thread thisThread = System.Threading.Thread.CurrentThread;
            thisThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        /// <summary>
        /// Get the column position based on its name (as string)
        /// </summary>
        /// <param name="strColumnName">teh column name (e.g. A, B, BA, ...)</param>
        /// <returns>The column number</returns>
        /// <example>exelExportHelper.getColumnnumber("C") will return 3</example>
        /// <remarks></remarks>
        public static int GetColumnNumber(string strColumnName)
        {
            int functionReturnValue = 0;
            //
            // Return column number based on column name
            // example column Z --> return 26 (Z is the 26 th column)
            // ------------------------------------------------------
            if (strColumnName.Length == 1)
            {
                return GetColumnNumber(strColumnName[0]);

            }
            else if (strColumnName.Length == 2)
            {
                functionReturnValue = (26 * GetColumnNumber(strColumnName[0])) + GetColumnNumber(strColumnName[1]);
            }
            else
            {
                throw new Exception("Column name is invalid : " + strColumnName);
            }
            return functionReturnValue;

        }

        /// <summary>
        /// Get the column position based on its name (as character)
        /// </summary>
        /// <param name="chrColumnName"></param>
        /// <returns></returns>
        public static int GetColumnNumber(char chrColumnName)
        {
            return (int)chrColumnName - 64;
        }

        /// <summary>
        /// changes the culture back to the one passed as parameter
        /// </summary>
        /// <param name="originalCulture">The new culture</param>
        /// <remarks></remarks>
        public static void RestoreCulture(System.Globalization.CultureInfo originalCulture)
        {
            //
            // Restore original culture of the user (necessary after change of the culture for the opening of the XLS)
            // ------------------------------------------------------------------------------------------------------- 
            System.Threading.Thread thisThread = System.Threading.Thread.CurrentThread;

            thisThread.CurrentCulture = originalCulture;
        }

        #endregion

        #endregion
    }
}
