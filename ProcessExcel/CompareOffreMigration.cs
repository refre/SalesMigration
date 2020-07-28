using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using System;

namespace Ista.Migration.Excel
{
    public class CompareOffreMigration
    {
        /// <summary>
        /// Private. migration variable
        /// </summary>
        private IList<string> _migration;

        /// <summary>
        /// Private. tbleoffre variable
        /// </summary>
        private DataTable _tblOffre;

        /// <summary>
        /// Private. List of offers.
        /// </summary>
        private List<DataTable> _offres;

        /// <summary>
        /// Private. contenu Variable.
        /// </summary>
        private List<BuildingFound> _contenu;

        /// <summary>
        /// Gets the number of counter.
        /// </summary>
	    public int Count
	    {
            get { return _contenu.Count; }
	    }

        /// <summary>
        /// Gets the list of 
        /// </summary>
        public List<BuildingFound> Contenu
        {
            get{return _contenu;}
        }

        /// <summary>
        /// Initializa a new CompareOffreMigration
        /// </summary>
        /// <param name="migration">List of migration element.</param>
        /// <param name="offres">List of offer datatable.</param>
        public CompareOffreMigration(IList<string> migration,List<DataTable> offres)
        {
            _migration = migration;
            _offres = offres;
            _contenu = new List<BuildingFound>();

            Parallel.ForEach(_offres, currentTable => CompareOffre(_migration, currentTable));   
        }

        /// <summary>
        /// Method that compare offre.
        /// </summary>
        /// <param name="migration">List of migration element.</param>
        /// <param name="tblOffre">Offer datatable.</param>
        public void CompareOffre(IList<string> migration, DataTable tblOffre)
        {
            _migration = migration;
            //Test purpose
            //_migration.Add("009382");
            //_migration.Add("003384");
            _tblOffre = tblOffre;

            foreach (DataRow item in tblOffre.Rows)
            {
                if (migration.Contains(item[3].ToString()))
                {
                    Console.WriteLine(item[3].ToString() + "-" + item[27].ToString() + "-" + item[28].ToString());
                    _contenu.Add(new BuildingFound{Building=item[3].ToString(),OffreNumber=item[27].ToString(),DateSent = item[28].ToString()});
                }
            }
        }
    }
}
