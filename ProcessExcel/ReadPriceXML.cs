using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Ista.Migration.Excel
{
    /// <summary>
    /// This class is used in order to read the prices coming from the XLM template made.
    /// </summary>
    public class ReadPriceXML
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
        /// Method that read the price of an xml file.
        /// </summary>
        /// <param name="fileName"></param>
        public ReadPriceXML(string fileName)
        {
            _fileName = fileName;
            _xdoc = XDocument.Load(_fileName);
        }

        /// <summary>
        /// This method return a list of Tarifs.
        /// </summary>
        /// <param name="element">Migration element.</param>
        /// <param name="id">Id of the element.</param>
        /// <param name="quantite">Quantity of the element.</param>
        /// <returns>List of tarifs.</returns>
        public List<Tarifs> GetTarifsForImmeuble(MigrationElement element, string id, string quantite)
        {
            var returnedValues = from xml in _xdoc.Descendants("Articles")
                          where (id.Equals(xml.Attribute("ID").Value)) 
                          select new Tarifs
                          {
                              Id = xml.Attribute("ID").Value,
                              Name = xml.Attribute("nom").Value,
                              Releve =  float.Parse(xml.Attribute("releve").Value),
                              Value = float.Parse(xml.Attribute("value").Value),
                              Vente = float.Parse(xml.Attribute("vente").Value),
                              LocationMax15 = float.Parse(xml.Attribute("locationMax15").Value)
                          };
            return returnedValues.ToList();
        }

        /// <summary>
        /// This method return a list of tarifs for the element.
        /// </summary>
        /// <returns>List of tarifs.</returns>
        public List<Tarifs> GetTarifsForImmeuble()
        {
            var returnedValues = from xml in _xdoc.Descendants("Articles")                          
                                 select new Tarifs
                                    {
                                        Id = xml.Attribute("ID").Value,
                                        Name = xml.Attribute("nom").Value,
                                        Releve = float.Parse(xml.Attribute("releve").Value),
                                        Value = float.Parse(xml.Attribute("value").Value),
                                        Vente = float.Parse(xml.Attribute("vente").Value),
                                        LocationMax15 = float.Parse(xml.Attribute("locationMax15").Value)
                                    };
            return returnedValues.ToList();
        }
    }
}
