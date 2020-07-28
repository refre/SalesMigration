using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.ComponentModel;
using System.Collections.ObjectModel;


namespace Ista.Migration.Excel
{
    /// <summary>
    /// This class is used in order format data and be able to display it in a grid.
    /// Code comes form:
    ///<code>
    /// http://www.c-sharpcorner.com/uploadfile/mgold/importing-an-excel-file-into-a-silverlight-datagrid-in-xml-format/
    /// </code>
    /// </summary>
    public class XmlGridProvider
    {
        /// <summary>
        /// Private field: filename. 
        /// </summary>
        private string _fileName;

        /// <summary>
        /// Private field: xml document. 
        /// </summary>
        private XDocument _xdoc;

        /// <summary>
        /// Gets or sets the filename.
        /// </summary>
        public string FileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }

        /// <summary>
        /// Initialize XmlGridProvider.
        /// </summary>
        /// <param name="fileName"></param>
        public XmlGridProvider (string fileName)
        {
            _fileName = fileName;
            _xdoc = XDocument.Load(_fileName);
        }

        /// <summary>
        /// Method to get column name
        /// </summary>
        /// <returns></returns>
        public List<string> GetColumnNameList()
        {
            // Don't forget the skip that allow to skip the first row of xml/Excel file
            return _xdoc.Descendants().Where(x => x.Name.LocalName == "Row").Skip(1).First().Descendants().Where(y => y.Name.LocalName == "Data").Select(q => q.Value).ToList();
        }

        /// <summary>
        /// Create list with data from xml doc.
        /// </summary>
        /// <returns></returns>
        public List<List<string>> PopulateData()
        {
            List<List<string>> items = new List<List<string>>();
            
            var popules = _xdoc.Descendants().Where(x => x.Name.LocalName == "Row").Skip(2);

            foreach (var item in popules)
            {
                    var data = item.Descendants().Where(y => y.Name.LocalName == "Data").Select(q => q.Value).ToList();
                    items.Add(data); 
            }   
            return items;
        }

        /// <summary>
        /// Create a list of migration element where the city = param cityIndex
        /// </summary>
        /// <param name="values"></param>
        /// <param name="cityIndex"></param>
        /// <returns></returns>
        public IList<MigrationElement> GetValueByCity(ObservableCollection<MigrationElement> values, string cityIndex)
        {
            var cityList = from c in values
                       where c.Site == cityIndex
                       select c;

            return cityList.ToList();
        }

        /// <summary>
        /// Create List of migration element to populate the grid with good value in good field.
        /// </summary>
        /// <param name="basePath"></param>
        /// <returns></returns>
        public List<MigrationElement> GetGridData(string basePath)
        {
            List<MigrationElement> data = new List<MigrationElement>();

            string[] antw = new string[] { "2", "8", "90", "91", "92", "97", "98", "99", "35", "36", "37", "39" };
            string[] brux = new string[] { "30", "31", "32", "33", "34", "38", "1", "7", "60", "61", "62", "64", "93", "94", "95", "96" };
            string[] Verv = new string[] { "4", "5", "66", "67", "69" };

            foreach (var item in PopulateData())
            {
                MigrationElement current      = new MigrationElement();
                current.PhysicalPerson        = false;
                current.Site                  = item[0];
                current.NumeroDeGroupe        = item[1];
                current.NomDeGroupe           = item[2];
                current.AdresseGroupe1        = item[3];
                current.AdresseGroupe2        = item[4];
                current.AdresseGroupe3        = item[5];
                current.CodePostauxGroupe     = item[6];
                current.LocaliteGroupe        = item[7];
                current.NumeroImmeuble        = item[8];
                current.NomImmeuble1          = item[9];
                current.NomImmeuble2          = item[10];
                current.AdresseImmeuble1      = item[11];
                current.AdresseImmeuble2      = item[12];
                current.AdresseImmeuble3      = item[13];
                current.CodePostalImmeuble    = item[14];
                current.LocaliteImmeuble      = item[15];
                current.NumeroDuGerant        = item[16];
                current.NomDeGerant           = item[17];
                current.FJuridique            = item[18];
                current.NomDeGerant2          = item[19];
                current.AdresseGerant1        = item[20];
                current.AdresseGerant2        = item[21];
                current.AdresseGerant3        = item[22];
                current.CodePostauxGerant     = item[23];
                current.LocaliteGerant        = item[24];
                current.Langue                = item[25];
                current.NombreAppartement     = item[26];
                current.ChauffageType         = item[27];
                current.ChauffageDescr        = item[28];
                // 27/05/2013 ajout de nouvelles colonnes dans ISEC
                // pusqu'il y a des colonne prix
                current.ChauffageNombreRad    = item[30];
                // ATTENTION RAJOUT D'UNE COLONNE DANS LE FICHIER XML(total pour doprimo)
                // => +1 au item ci dessous
                current.ChauffageNombreNRad   = item[31];
                current.EauChaudeType         = item[32];
                current.EauChaudeDescr        = item[33];
                // 27/05/2013 ajout de nouvelles colonnes dans ISEC
                // pusqu'il y a des colonne prix
                current.EauChaudeNombreRad    = item[35];
                current.EauChaudeNombreNRad   = item[36];
                current.EauFroideType         = item[37];
                current.EauFroideDescr        = item[38];
                // 27/05/2013 ajout de nouvelles colonnes dans ISEC
                // pusqu'il y a des colonne prix
                current.EauFroideNombreRad    = item[40];
                current.EauFroideNombreNRad   = item[41];
                current.IntegrateurType       = item[42];
                current.IntegrateurDescr      = item[43];
                // 27/05/2013 ajout de nouvelles colonnes dans ISEC
                // pusqu'il y a des colonne prix
                current.IntegrateurNombreRad  = item[45];
                current.IntegrateurNombreNRad = item[46];

                if (current.CodePostalImmeuble.StartsWithAny(Verv))
                {
                    current.DocumentName = "Migr-VV-" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.ToString("MM") + "-" + current.NumeroImmeuble;
                }
                else if (current.CodePostalImmeuble.StartsWithAny(antw))
                {
                    current.DocumentName = "Migr-AN-" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.ToString("MM") + "-" + current.NumeroImmeuble;
                }
                else
                {
                    current.DocumentName = "Migr-BX-" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.ToString("MM") + "-" + current.NumeroImmeuble;
                }
                current.Element = new WordElement(current, basePath);
                
                data.Add(current);
            }

            return data;
        }
        
        /// <summary>
        /// method to return a list of column name from the index
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public List<string> GetAnyColumn(int columnIndex)
        {
            List<string> column = new List<string>();

            foreach (var item in PopulateData())
            {
                column.Add(item[columnIndex]);
            }
            return column;
        }

        /// <summary>
        /// method to get numbers of group
        /// </summary>
        /// <returns></returns>
        public List<string> GetGroupNumbers()
        {
            return GetAnyColumn(1);
        }

        /// <summary>
        /// method to get group name
        /// </summary>
        /// <returns></returns>
        public List<string> GetGroupName()
        {
            return GetAnyColumn(2);
        }

        /// <summary>
        /// method to get the group adress
        /// </summary>
        /// <returns></returns>
        public List<string> GetGroupAddress()
        {
            return GetAnyColumn(3);
        }

        /// <summary>
        /// method to get number of building.
        /// </summary>
        /// <returns></returns>
        public List<string> GetImmeubleNumber()
        {
            return GetAnyColumn(8);
        }
    }
}
