using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Data;

namespace Ista.Migration.Excel
{
    /// <summary>
    /// This class manages a connection toward an Excel 2007 file.
    /// It uses an OLEDB connection.
    /// http://www.daniweb.com/software-development/csharp/threads/130123
    /// </summary>
    public class Excel12Manager
    {
        /// <summary>
        /// Portected variable connection string builder.
        /// </summary>
        protected OleDbConnectionStringBuilder _sbConnection;

        /// <summary>
        /// Portected variable Path of the file to process.
        /// </summary>
        protected string _path;

        /// <summary>
        /// Gets the Path of the Excel File.
        /// </summary>
        public string PathFileName
        {
            get
            {
                return _path;
            }
        }

        /// <summary>
        /// Gets the connection string to the Excel 2007 file.
        /// </summary>
        public string ConnectionString
        {
            get
            {
                return _sbConnection.ToString();
            }
        }

        /// <summary>
        /// Initializes a new Excel12Manager.
        /// </summary>
        /// <param name="path">Path of the Excel 2007 file.</param>
        public Excel12Manager(string path)
        {
            _path = path;
            
            string strExtendedProperties = string.Empty;
            _sbConnection = new OleDbConnectionStringBuilder();

            if (Path.GetExtension(_path).Equals(".xlsx"))
            {
                _sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
                _sbConnection.DataSource = _path;
                strExtendedProperties = "Excel 12.0;HDR=Yes;IMEX=1";
            }
            else 
            {
                throw new Exception("");
            }
            _sbConnection.Add("Extended Properties",strExtendedProperties);

        }

        /// <summary>
        /// Gets the several worksheet names in the ocument
        /// </summary>
        /// <returns>List of the worksheetName</returns>
        public IList<string> GetWorkSheetName()
        {
            List<string> listSheet = new List<string>();
            using (OleDbConnection conn = new OleDbConnection(_sbConnection.ToString()))
            {
                conn.Open();
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // TODO: Check if dtSheet is null.
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                    //checks whether row contains '_xlnm#_FilterDatabase' or sheet name(i.e. sheet name always ends with $ sign)         
                    {
                        string currentName = drSheet["TABLE_NAME"].ToString();
                        string DispalyName = currentName.Remove(currentName.Length - 1);
                        listSheet.Add(DispalyName);
                    }
                }
            }
            return listSheet;
        }

        public virtual DataTable ReadWorkSheet(string workSheetName)
        {
            return null;
        }
    }
}
