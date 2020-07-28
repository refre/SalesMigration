using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using Ista.Migration.Excel;
using Microsoft.Win32;
using System.Threading;
using System.Globalization;
using System.Linq;


namespace Ista.Migration
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Variable

        /// <summary>
        /// Private. Migration path
        /// </summary>
        private string _migrationPath;
        /// <summary>
        /// Private Offre Bx path;
        /// </summary>
        private string _offerBxPath;
        /// <summary>
        /// Private offre AN Path
        /// </summary>
        private string _offerAnPath;
        /// <summary>
        /// Private offre VV Path
        /// </summary>
        private string _offerVvPath;
        private const string _fileName = "Migration.exe";
        private readonly string _filePath;
        private readonly string _basePath;
        private List<MigrationElement> _valFr;
        private List<MigrationElement> _valNl;

        #endregion Variable

        #region Constructor

        /// <summary>
        /// Initialize the main window.
        /// Put currentculture to fr-BE to be sure to have coma in price
        /// </summary>
        public MainWindow()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-BE");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-BE");
            InitializeComponent();

            ComapreEvent += new ComapreEventHandler(ComapreEnableEvent);

            _filePath = Assembly.GetExecutingAssembly().Location;
            _basePath = _filePath.Remove(_filePath.Length - _fileName.Length, _fileName.Length);


        }

        #endregion Constructor

        #region Private method

        /// <summary>
        /// Check whether the data are enable in order to make the comparaison enable.
        /// </summary>
        private void ComapreEnableEvent()
        {
            bool offre = false;
            bool migration = txtMigration.Text.Length > 0;

            if (txtOffre.Text.Length > 0 | txtOffreAN.Text.Length > 0 | txtOffreVV.Text.Length > 0)
                offre = true;

            if (migration && offre)
                btnCompare.IsEnabled = true;
            else
                btnCompare.IsEnabled = false;
        }

        /// <summary>
        /// clear txtCompareResult
        /// </summary>
        private void clearAll()
        {
            txtCompareResult.Text = "";
        }

        // Wrap the event in a protected virtual method
        // to enable derived classes to raise the event.
        protected virtual void ValidateComapre()
        {
            // Raise the event by using the () operator.
            ComapreEvent();
        }

        /// <summary>
        /// Method to call CreateWordDocument method
        /// </summary>
        /// <param name="currentElement"></param>
        private void Process(MigrationElement currentElement)
        {
            try
            {
                string fileIn;
                WordReadWrite WordProcess = new WordReadWrite();

                if (!Directory.Exists(_basePath + Properties.Settings.Default.ResultPath))
                {
                    Directory.CreateDirectory(_basePath + Properties.Settings.Default.ResultPath);
                }

                string pathOut = _basePath + Properties.Settings.Default.ResultPath;
                string pathPhyFr = _basePath + Properties.Settings.Default.PathPhysFrTemplate;
                string pathPhyNl = _basePath + Properties.Settings.Default.PathPhysNlTemplate;
                string pathMorFr = _basePath + Properties.Settings.Default.PathMoraFrTemplate;
                string pathMorNl = _basePath + Properties.Settings.Default.PathMoraNlTemplate;

                if (currentElement.PhysicalPerson)
                {
                    fileIn = currentElement.Langue.Trim().ToUpper().Equals("N") ? pathPhyNl : pathPhyFr;
                }
                else
                {
                    fileIn = currentElement.Langue.Trim().ToUpper().Equals("N") ? pathMorNl : pathMorFr;
                }

                string fileOut;
                if (currentElement.PhysicalPerson)
                    fileOut = pathOut + currentElement.DocumentName + ".docx";
                else
                    fileOut = pathOut + currentElement.DocumentName + ".docx";

                WordProcess.CreatWordDocument(fileIn, fileOut, currentElement);

                System.Diagnostics.Process.Start(pathOut);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion

        #region Event method

        /// <summary>
        /// Event method to open Xml File
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMigration_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".xml"; // Default file extension
            dlg.Filter = "xml documents (.xml)|*.xml"; // Filter files by extension
            bool? result = dlg.ShowDialog();
            if (result==true)
            {
                _migrationPath = dlg.FileName;
                txtMigration.Text = _migrationPath;
            }
            ComapreEvent();
        }

        /// <summary>
        /// Event method to open Excel File
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOffre_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel 2007-2010 documents (.xlsx)|*.xlsx"; // Filter files by extension
            bool? result = dlg.ShowDialog();
            if (result==true)
            {
                if (e.OriginalSource.Equals(this.btnOffre))
                {
                    _offerBxPath = dlg.FileName;
                    txtOffre.Text = _offerBxPath;
                }
                else if (e.OriginalSource.Equals(this.btnLoadAN))
                {
                    _offerAnPath = dlg.FileName;
                    txtOffreAN.Text = _offerAnPath;
                }
                else if (e.OriginalSource.Equals(this.btnLoadVV))
                {
                    _offerVvPath = dlg.FileName;
                    txtOffreVV.Text = _offerVvPath;
                }
            }
            ComapreEvent();
        }

        /// <summary>
        /// Occurs when the user clicks on the btn Comapre.
        /// </summary>
        /// <exception cref="http://refactormycode.com/codes/343-merge-generic-lists">Code in order to merge list.</exception>
        private void btnCompare_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            try
            {
                List<string> filePath = new List<string>();
                XmlGridProvider gridDataProvider;
                //ObservableCollection<MigrationElement> migrationDisplay = new ObservableCollection<MigrationElement>();
                List<MigrationElement> migrationDisplay = new List<MigrationElement>();

                if (!string.IsNullOrEmpty(_offerBxPath))
                {
                    filePath.Add(_offerBxPath);
                }
                if (!string.IsNullOrEmpty(_offerAnPath))
                {
                    filePath.Add(_offerAnPath);
                }
                if (!string.IsNullOrEmpty(_offerVvPath))
                {
                    filePath.Add(_offerVvPath);
                }

                StringBuilder comparedvaluebuilder = new StringBuilder();
                string comparedValue = "";

                List<MigrationElement> currentElement = new List<MigrationElement>();
                gridDataProvider = new XmlGridProvider(_migrationPath);
                var val = gridDataProvider.GetGridData(_basePath + "XML\\tarifsCopy2.xml");

                foreach (var item in filePath)
                {
                    var groupNumbers = gridDataProvider.GetGroupNumbers();
                    OffresExcel offreFiles = new OffresExcel(item);
                    offreFiles.GetWorkSheetName();
                    List<DataTable> tblOffres = offreFiles.GetCompleteExcelInList();

                    CompareOffreMigration myCompare = new CompareOffreMigration(groupNumbers, tblOffres);

                    if (myCompare.Count == 0)
                    {
                        if (comparedvaluebuilder.ToString() == string.Empty)
                            comparedValue = "There is no contract alredy sent.";
                    }
                    else
                    {
                        foreach (BuildingFound building in myCompare.Contenu)
                        {
                            comparedvaluebuilder.Append("The contract for the bulding ");
                            comparedvaluebuilder.Append(building.Building);
                            comparedvaluebuilder.Append(" was sent ont the ");
                            comparedvaluebuilder.Append(building.DateSent);
                            comparedvaluebuilder.Append(" with the offer ");
                            comparedvaluebuilder.Append(building.OffreNumber);
                            comparedvaluebuilder.Append(" \n");
                        }
                        comparedValue = comparedvaluebuilder.ToString();
                    }

                    foreach (var item2 in myCompare.Contenu)
                    {
                        // Search where the building match then remove it.
                        // var migElement = val.Find(x => x.NumeroDeGroupe == item2.Building);
                        var migElement = val.Find(x => x.NumeroImmeuble == item2.Building);
                        if (item != null)
                            val.Remove(migElement);
                    }
                }

                dataMigration.ItemsSource = val;

                _valNl = val.Where(x => x.Langue.ToUpper() == "N").ToList();
                _valFr = val.Where(x => x.Langue.ToUpper() != "N").ToList();

                txtCompareResult.Text = comparedValue;
                this.Cursor = Cursors.Arrow;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        /// <summary>
        /// put data in Item price panel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataMigration_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataMigration.SelectedItem == null)
                return;
            try
            {
                _currentElement = (MigrationElement)dataMigration.SelectedItem;

                txtDoprimoIIIRQuant.Text = _currentElement.ChauffageNombreNRad;
                txtDoprimoIIIRPrice.Text = _currentElement.Element.Dop3RadioChaufVente;
                txtDoprimoIIIRRent.Text = _currentElement.Element.Dop3RadioChaufLocat;
                txtDoprimoIIIRRead.Text = _currentElement.Element.Dop3RadioChaufRelev;

                txtDomaquaRQuant.Text = _currentElement.EauFroideNombreNRad;
                txtDomaquaRPrice.Text = _currentElement.Element.DomaAquaEauFrRVente;
                txtDomaquaRRent.Text = _currentElement.Element.DomaAquaEauFrRLocat;
                txtDomaquaRRead.Text = _currentElement.Element.DomaAquaEauFrRRelev;

                txtDomaquaRHotQuant.Text = _currentElement.EauChaudeNombreNRad;
                txtDomaquaRHotPrice.Text = _currentElement.Element.DomaAquaEauChRVente;
                txtDomaquaRHotRent.Text = _currentElement.Element.DomaAquaEauChRLocat;
                txtDomaquaRHotRead.Text = _currentElement.Element.DomaAquaEauChRRelev;

                txtSenso1_2RQuant.Text = _currentElement.IntegrateurNombreNRad;
                txtSenso1_2RPrice.Text = _currentElement.Element.SensonicIn1_2Vente;
                txtSenso1_2RRent.Text = _currentElement.Element.SensonicIn1_2Locat;
                txtSenso1_2RRead.Text = _currentElement.Element.SensonicIn1_2Relev;

                txtSensoR3_4Quant.Text = _currentElement.IntegrateurNombreNRad;
                txtSensoR3_4Price.Text = _currentElement.Element.SensonicIn3_4Vente;
                txtSensoR3_4Rent.Text = _currentElement.Element.SensonicIn3_4Locat;
                txtSensoR3_4Read.Text = _currentElement.Element.SensonicIn3_4Relev;

                txtDomaquaTotQuant.Text = _currentElement.Element.DomaquaTotal;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Empty cell: just click to in a cell with data");
            }
        }
        
        /// <summary>
        /// put the data of item price in element of grid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSavePrices_Click(object sender, RoutedEventArgs e)
        {
            _currentElement.ChauffageNombreNRad         = txtDoprimoIIIRQuant.Text;
            _currentElement.Element.Dop3RadioChaufVente = txtDoprimoIIIRPrice.Text;
            _currentElement.Element.Dop3RadioChaufLocat = txtDoprimoIIIRRent.Text;
            _currentElement.Element.Dop3RadioChaufRelev = txtDoprimoIIIRRead.Text;

            _currentElement.EauFroideNombreNRad         = txtDomaquaRQuant.Text;
            _currentElement.Element.DomaAquaEauFrRVente = txtDomaquaRPrice.Text;
            _currentElement.Element.DomaAquaEauFrRLocat = txtDomaquaRRent.Text;
            _currentElement.Element.DomaAquaEauFrRRelev = txtDomaquaRRead.Text;

            _currentElement.EauChaudeNombreNRad         = txtDomaquaRHotQuant.Text;
            _currentElement.Element.DomaAquaEauChRVente = txtDomaquaRHotPrice.Text;
            _currentElement.Element.DomaAquaEauChRLocat = txtDomaquaRHotRent.Text;
            _currentElement.Element.DomaAquaEauChRRelev = txtDomaquaRHotRead.Text;

            _currentElement.IntegrateurNombreNRad       = txtSenso1_2RQuant.Text;
            _currentElement.Element.SensonicIn1_2Vente  = txtSenso1_2RPrice.Text;
            _currentElement.Element.SensonicIn1_2Locat  = txtSenso1_2RRent.Text;
            _currentElement.Element.SensonicIn1_2Relev  = txtSenso1_2RRead.Text;

            _currentElement.IntegrateurNombreNRad       = txtSensoR3_4Quant.Text;
            _currentElement.Element.SensonicIn3_4Vente  = txtSensoR3_4Price.Text;
            _currentElement.Element.SensonicIn3_4Locat  = txtSensoR3_4Rent.Text;
            _currentElement.Element.SensonicIn3_4Relev  = txtSensoR3_4Read.Text;


        }
        /// <summary>
        /// This method checks whether there are already some excel process running.
        /// </summary>
        private void CheckWordProcesses()
        {
            Process[] AllProcesses = System.Diagnostics.Process.GetProcessesByName("winword");
            _myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process WordProcess in AllProcesses)
            {
                _myHashtable.Add(WordProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }

        /// <summary>
        /// This method kills excel process.
        /// </summary>
        private void KillWord()
        {
            Process[] AllProcesses = System.Diagnostics.Process.GetProcessesByName("winword");
            // check to kill the right process
            foreach (Process wordProcess in AllProcesses)
            {
                if (_myHashtable.ContainsKey(wordProcess.Id) == false)
                    wordProcess.Kill();
            }
            AllProcesses = null;
        }
        
        /// <summary>
        /// The same method of btnWord Click but for each record
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAll_Click(object sender, RoutedEventArgs e)
        {
            if (dataMigration.Items.Count == 0)
                return;

            

            
            _myHashtable = new Hashtable();
            int iCount = 0;

            foreach (var item in _valFr)
            {
                
                Process(item);
                Process[] AllProcesses = System.Diagnostics.Process.GetProcessesByName("winword");
                foreach (Process WordProcess in AllProcesses)
                {
                    _myHashtable.Add(WordProcess.Id, iCount);
                    iCount = iCount + 1;
                }
                System.Threading.Thread.Sleep(3000);
                List<MigrationElement> element = new List<MigrationElement>();
                element.Add(item);

                foreach (Process wordProcess in AllProcesses)
                {
                    if (_myHashtable.ContainsKey(wordProcess.Id) == true)
                        wordProcess.Kill();
                }
                ReportCreator myCreator;
                try
                {
                    myCreator = new ReportCreator(element, _offerBxPath, _offerAnPath, _offerVvPath);
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    AllProcesses = null;
                    element = null;
                    myCreator = null;
                }
            }

            System.Threading.Thread.Sleep(500);

            foreach (var item in _valNl)
            {

                Process(item);
                Process[] AllProcesses = System.Diagnostics.Process.GetProcessesByName("winword");
                foreach (Process WordProcess in AllProcesses)
                {
                    _myHashtable.Add(WordProcess.Id, iCount);
                    iCount = iCount + 1;
                }
                System.Threading.Thread.Sleep(3000);
                List<MigrationElement> element = new List<MigrationElement>();
                element.Add(item);

                foreach (Process wordProcess in AllProcesses)
                {
                    if (_myHashtable.ContainsKey(wordProcess.Id) == true)
                        wordProcess.Kill();
                }
                ReportCreator myCreator;
                try
                {
                    myCreator = new ReportCreator(element, _offerBxPath, _offerAnPath, _offerVvPath);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    AllProcesses = null;
                    element = null;
                    myCreator = null;
                }
            }

            //CheckWordProcesses();
            //foreach (var item in dataMigration.ItemsSource)
            //{

            //    foreach (Process WordProcess in AllProcesses)
            //    {
            //        _myHashtable.Add(WordProcess.Id, iCount);
            //        iCount = iCount + 1;
            //    }

            //    MigrationElement currentElement = (MigrationElement)item;
            //    Process(currentElement);
            //    System.Threading.Thread.Sleep(1000);
            //    element.Add(currentElement);

            //    foreach (Process wordProcess in AllProcesses)
            //    {
            //        if (_myHashtable.ContainsKey(wordProcess.Id) == false)
            //            wordProcess.Kill();
            //    }
            //}
            //KillWord();
           

            MessageBox.Show(string.Format("The files are available."));
        }

        /// <summary>
        /// Event method to call method to create doc word and call method to modify the excel file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWord_Click(object sender, RoutedEventArgs e)
        {
            if (dataMigration.SelectedItem == null)
                return;

            CheckWordProcesses();

            MigrationElement currentElement = (MigrationElement)dataMigration.SelectedItem;
            Process(currentElement);

            List<MigrationElement> element = new List<MigrationElement>();
            element.Add(currentElement);

            KillWord();
            try
            {
                ReportCreator myCreator = new ReportCreator(element, _offerBxPath, _offerAnPath, _offerVvPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// Event method to close application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Event method to close application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Event method to show a new windows with price
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            Prices price = new Prices();
            price.ShowDialog();
        }

        #endregion
        
        private MigrationElement _currentElement; 
       
        public delegate void ComapreEventHandler();

        // Declare the event.
        public event ComapreEventHandler ComapreEvent;
        
        /// <summary>
        /// Private variable: Hastable of the excel process
        /// </summary>
        private Hashtable _myHashtable;
    }
}
