using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Reflection;

namespace Ista.Migration
{
    /// <summary>
    /// Interaction logic for Prices.xaml
    /// </summary>
    public partial class Prices : Window
    {
        private const string _fileName = "Migration.exe";
        private readonly string _filePath;
        private string _basePath;

        public Prices()
        {
            InitializeComponent();
            _filePath = Assembly.GetExecutingAssembly().Location;
            _basePath = _filePath.Remove(_filePath.Length - _fileName.Length, _fileName.Length);   
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            /* Source="XML\tarifs.xml"*/
            XmlDataProvider xmlfile = (XmlDataProvider)FindResource("xmlfile");
            xmlfile.Source = new Uri(_basePath + "\\XML\\tarifsCopy2.xml");
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            XmlDataProvider xmlfile = (XmlDataProvider)FindResource("xmlfile");

            string source = xmlfile.Source.LocalPath;
            xmlfile.Document.Save(source);
        }
    }
}
