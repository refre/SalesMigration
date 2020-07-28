using System.ComponentModel;

namespace Ista.Migration.Excel
{
    /// <summary>
    /// This class is the object containign the migration elements.
    /// </summary>
    public class MigrationElement : INotifyPropertyChanged
    {
        /// <summary>
        /// Private variable: physical person.
        /// </summary>
        private bool _physicalPerson;

        /// <summary>
        /// Gets or sets whether the contract id the physical person.
        /// </summary>
        public bool PhysicalPerson
        {
            get { return _physicalPerson; } 
            set
            {
                if (value != _physicalPerson)
                {
                    _physicalPerson = value;
                    OnPropertyChanged("PhysicalPerson");
                }
            }
        }

        /// <summary>
        /// Gets or sets the site name.
        /// </summary>
        public string Site                  { get; set; }

        /// <summary>
        /// Gets or sets the "Numero de groupe".
        /// </summary>
        public string NumeroDeGroupe        { get; set; }

        /// <summary>
        /// Gets or sets the "Nom de groupe".
        /// </summary>
        public string NomDeGroupe           { get; set; }

        /// <summary>
        /// Gets or sets the "Adresse groupe 1".
        /// </summary>
        public string AdresseGroupe1        { get; set; }

        /// <summary>
        /// Gets or sets the "Adresse groupe 2".
        /// </summary>
        public string AdresseGroupe2        { get; set; }

        /// <summary>
        /// Gets or sets the "Adresse groupe 3".
        /// </summary>
        public string AdresseGroupe3        { get; set; }

        /// <summary>
        /// Gets or sets the "Code Postaux groupe".
        /// </summary>
        public string CodePostauxGroupe     { get; set; }

        /// <summary>
        /// Gets or sets the "Localité groupe".
        /// </summary>
        public string LocaliteGroupe        { get; set; }
        public string NumeroImmeuble        { get; set; }
        public string NomImmeuble1          { get; set; }
        public string NomImmeuble2          { get; set; }
        public string AdresseImmeuble1      { get; set; }
        public string AdresseImmeuble2      { get; set; }
        public string AdresseImmeuble3      { get; set; }
        public string CodePostalImmeuble    { get; set; }
        public string LocaliteImmeuble      { get; set; }
        public string NumeroDuGerant        { get; set; }
        public string NomDeGerant           { get; set; }
        public string NomDeGerant2          { get; set; }
        public string AdresseGerant1        { get; set; }
        public string AdresseGerant2        { get; set; }
        public string AdresseGerant3        { get; set; }
        public string CodePostauxGerant     { get; set; }
        public string LocaliteGerant        { get; set; }
        public string Langue                { get; set; }
        public string ChauffageType         { get; set; }
        public string ChauffageDescr        { get; set; }
        public string ChauffageNombreRad    { get; set; }
        public string ChauffageNombreNRad   { get; set; }

        public string EauChaudeType         { get; set; }
        public string EauChaudeDescr        { get; set; }
        public string EauChaudeNombreRad    { get; set; }
        public string EauChaudeNombreNRad   { get; set; }

        public string EauFroideType         { get; set; }
        public string EauFroideDescr        { get; set; }
        public string EauFroideNombreRad    { get; set; }
        public string EauFroideNombreNRad   { get; set; }

        public string IntegrateurType       { get; set; }
        public string IntegrateurDescr      { get; set; }
        public string IntegrateurNombreRad  { get; set; }
        public string IntegrateurNombreNRad { get; set; }

        public string DiversType            { get; set; }
        public string DiversDescr           { get; set; }
        public string DiversNombreRad       { get; set; }
        public string DiversNombreNRad      { get; set; }

        public string FJuridique            { get; set; }
        public string NombreAppartement     { get; set; }
        public string DocumentName          { get; set; }
        public int NumberOfRows             { get; set; }

        public WordElement Element          { get; set; }
        


        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        
    }
}
