using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ista.Migration.Excel
{

    public enum ProductsIsta
    {
        doprimoIIIRadio,
        doprimoIIIRs,
        domaquaRadio,
        sensonicRadio1_2,
        sensonicRadio3_4
    };

    /// <summary>
    /// This class is the BO that manages calcutaion of any price regarding quantities.
    /// </summary>
    public class WordElement
    {

        #region Properties
        /// <summary>
        /// Gets or sets Code name of the gerant.
        /// </summary>
        public string CodeNameGerant { get; private set; }
        /// <summary>
        /// Gets or sets Code adresse gerant.
        /// </summary>
        public string CodeAdresseGerant { get; private set; }
        /// <summary>
        /// Gets or sets Total adresse gerant
        /// </summary>
        public string TotalAdresseGerant { get; private set; }
        /// <summary>
        /// Gets or sets Code offre number.
        /// </summary>
        public string CodeOffreNum { get; private set; }
        /// <summary>
        /// Gets or sets Code immeuble
        /// </summary>
        public string CodeImmeuble { get; private set; }
        /// <summary>
        /// Gets or sets code addresse Immeuble.
        /// </summary>
        public string CodeAdressIm { get; private set; }
        /// <summary>
        /// Gets or sets the Total adresse immeuble.
        /// </summary>
        public string TotalAdressImeuble { get; set; }
        /// <summary>
        /// Gets or sets the NbrCahuf.
        /// </summary>
        public string NbrChauf { get; set; }
        /// <summary>
        /// Gets or sets the NbrNChauf.
        /// </summary>
        public string NbrNChauf { get; set; }
        /// <summary>
        /// Gets or sets the NbrEauCh.
        /// </summary>
        public string NbrEauCh { get; set; }
        /// <summary>
        /// Gets or sets the NbrNEauCh.
        /// </summary>
        public string NbrNEauCh { get; set; }
        /// <summary>
        /// Gets or sets the NbrEauFr.
        /// </summary>
        public string NbrEauFr { get; set; }
        /// <summary>
        /// Gets or sets NbrNEauFr.
        /// </summary>
        public string NbrNEauFr { get; set; }
        /// <summary>
        /// Gets or sets NbrInteg.
        /// </summary>
        public string NbrInteg { get; set; }
        /// <summary>
        /// Gets or sets NbrNInteg.
        /// </summary>
        public string NbrNInteg { get; set; }
        /// <summary>
        /// Gets or sets TotNchauf.
        /// </summary>
        public int TotNchauf { get; set; }
        /// <summary>
        /// Gets or sets TotNEauch.
        /// </summary>
        public int TotNEauch { get; set; }
        /// <summary>
        /// Gets or sets TotNEauFr.
        /// </summary>
        public int TotNEauFr { get; set; }
        /// <summary>
        /// Gets or sets TotInteg.
        /// </summary>
        public int TotInteg { get; set; }
        /// <summary>
        /// Gets or sets the total number of water counter.
        /// </summary>
        public int TotEau { get; set; }
        /// <summary>
        /// Gets or sets the total number of devices
        /// </summary>
        public int TotalDevice { get; set; }
        /// <summary>
        /// Gets or sets the chauf prices.
        /// </summary>
        public List<Tarifs> NChaufPrices { get; set; }
        /// <summary>
        /// Gets or sets the tarif list for hot water.
        /// </summary>
        public List<Tarifs> NEauChPrices { get; set; }
        /// <summary>
        /// Gets or sets the tarif list for cold water.
        /// </summary>
        public List<Tarifs> NEauFrPrices { get; set; }
        /// <summary>
        /// Gets or sets the tarif list for Intergators.
        /// </summary>
        public List<Tarifs> NIntegPrices { get; set; }
        /// <summary>
        /// Gets or sets the tarif list for the total of water counter.
        /// </summary>
        public List<Tarifs> NToEauPrices { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Doprimo3 Value.
        /// </summary>
        public float Dop3RadioChaufVenteValue { get; set; }
        /// <summary>
        /// Gets or sets the reading price for the product Doprimo3 Value.
        /// </summary>
        public float Dop3RadioChaufRelevValue { get; set; }
        /// <summary>
        /// Gets or sets the rent price for the product Doprimo3 Value.
        /// </summary>
        public float Dop3RadioChaufLocatValue { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Doprimo3.
        /// </summary>
        public string Dop3RadioChaufVente { get; set; }
        /// <summary>
        /// Gets or sets the reading price for the product Doprimo3.
        /// </summary>
        public string Dop3RadioChaufRelev { get; set; }
        /// <summary>
        /// Gets or sets the rent price for the product Doprimo3.
        /// </summary>
        public string Dop3RadioChaufLocat { get; set; }       
        /// <summary>
        /// Gets or sets the sell price for the product Domaqua (hot).
        /// </summary>
        public string DomaAquaEauChRVente { get; set; }
        /// <summary>
        /// Gets or sets the reading price for the product Domaqua (hot).
        /// </summary>
        public string DomaAquaEauChRRelev { get; set; }
        /// <summary>
        /// Gets or sets the rent price for the product Domaqua (hot).
        /// </summary>
        public string DomaAquaEauChRLocat { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Domaqua (cold).
        /// </summary>
        public string DomaAquaEauFrRVente { get; set; }
        /// <summary>
        /// Gets or sets the read price for the product Domaqua (cold).
        /// </summary>
        public string DomaAquaEauFrRRelev { get; set; }
        /// <summary>
        /// Gets or sets the rent price for the product Domaqua (cold).
        /// </summary>
        public string DomaAquaEauFrRLocat { get; set; }
        /// <summary>
        /// Gets or sets the read price for the product Sensonic 1/2Value.
        /// </summary>
        public float SensonicIn1_2RelevValue { get; set; }
        /// <summary>
        /// Gets or sets the rent price for the product Sensonic 1/2 Value.
        /// </summary>
        public float SensonicIn1_2LocatValue { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Sensonic 1/2 Value.
        /// </summary>
        public float SensonicIn1_2VenteValue { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Sensonic 3/4 Value.
        /// </summary>
        public float SensonicIn3_4VenteValue { get; set; }
        /// <summary>
        /// Gets or sets the read price for the product Sensonic 3/4 Value.
        /// </summary>
        public float SensonicIn3_4RelevValue { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Sensonic 3/4.
        /// </summary>
        public float SensonicIn3_4LocatValue { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Sensonic 1/2.
        /// </summary>
        public string SensonicIn1_2Vente { get; set; }
        /// <summary>
        /// Gets or sets the read price for the product Sensonic 1/2.
        /// </summary>
        public string SensonicIn1_2Relev { get; set; }
        /// <summary>
        /// Gets or sets the rent price for the product Sensonic 1/2.
        /// </summary>
        public string SensonicIn1_2Locat { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Sensonic 3/4.
        /// </summary>
        public string SensonicIn3_4Vente { get; set; }
        /// <summary>
        /// Gets or sets the read price for the product Sensonic 3/4.
        /// </summary>
        public string SensonicIn3_4Relev { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Sensonic 3/4.
        /// </summary>
        public string SensonicIn3_4Locat { get; set; }
        /// <summary>
        /// Gets or sets the DomaquaTotal value.
        /// </summary>
        public string DomaquaTotal { get; set; }

        /// <summary>
        ///  Gets or sets the read price for the product Domaqua Value.
        /// </summary>
        public float DomaquaTotalReleveValue { get; set; }
        /// <summary>
        /// Gets or sets the sell price for the product Domaqua Value.
        /// </summary>
        public float DomaquaTotalVenteValue { get; set; }
        /// <summary>
        /// Gets or sets the rent price for the product Domaqua Value.
        /// </summary>
        public float DomaquaTotalLocatiValue { get; set; }
        /// <summary>
        /// Gets or sets the read price for the product Domaqua.
        /// </summary>
        public string DomaquaTotalReleve { get; set; }
        /// <summary>
        ///  Gets or sets the sell price for the product Domaqua.
        /// </summary>        
        public string DomaquaTotalVente { get; set; }
        /// <summary>
        /// Gets or sets the rent price for the product Domaqua.
        /// </summary>        
        public string DomaquaTotalLocati { get; set; }
        #endregion Properties

        #region Constructor
        /// <summary>
        /// Initialize a new word element to process.
        /// </summary>
        /// <param name="currentElement">Current Element.</param>
        /// <param name="tarifsPath">Path of the </param>
        public WordElement(MigrationElement currentElement, string tarifsPath)
        {
            CodeNameGerant = !string.IsNullOrEmpty(currentElement.NomDeGerant)
                                        ? currentElement.NomDeGerant : currentElement.NomDeGerant2;
            CodeAdresseGerant = !string.IsNullOrEmpty(currentElement.AdresseGerant2)
                                ? currentElement.AdresseGerant2: currentElement.AdresseGerant1 ;
            TotalAdresseGerant = CodeAdresseGerant + ", " + currentElement.CodePostauxGerant + " " + currentElement.LocaliteGerant;
            CodeOffreNum = currentElement.Site + "-" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.ToString("MM") + "-" + currentElement.NumeroImmeuble;
            CodeImmeuble = currentElement.NumeroImmeuble;
            CodeAdressIm = !string.IsNullOrEmpty(currentElement.AdresseImmeuble1)
                                           ? currentElement.AdresseImmeuble1
                                           : currentElement.AdresseImmeuble2;
            TotalAdressImeuble = CodeAdressIm + ", " + currentElement.CodePostalImmeuble + " " +
                                        currentElement.LocaliteImmeuble;

            NbrChauf = currentElement.ChauffageNombreRad;
            NbrNChauf = currentElement.ChauffageNombreNRad;

            NbrEauCh = currentElement.EauChaudeNombreRad;
            NbrNEauCh = currentElement.EauChaudeNombreNRad;

            NbrEauFr = currentElement.EauFroideNombreRad;
            NbrNEauFr = currentElement.EauFroideNombreNRad;

            NbrInteg = currentElement.IntegrateurNombreRad;
            NbrNInteg = currentElement.IntegrateurNombreNRad;

            List<Tarifs> prices = new ReadPriceXML(tarifsPath).GetTarifsForImmeuble();

            int totNchauf = 0;
            int totNEauch = 0;
            int totNEauFr = 0;
            int totInteg = 0;
            int totalEau = 0;

            NChaufPrices = new List<Tarifs>();
            NEauChPrices = new List<Tarifs>();
            NEauFrPrices = new List<Tarifs>();
            NToEauPrices = new List<Tarifs>();
            NIntegPrices = new List<Tarifs>();

            if (!string.IsNullOrEmpty(NbrNChauf))
            {
                int.TryParse(NbrNChauf, out totNchauf);
                TotNchauf = totNchauf;
                NChaufPrices = Compteur(totNchauf, prices, "doprimoIIIRadio", false,true);
            }

            if (!string.IsNullOrEmpty(NbrNInteg))
            {
                int.TryParse(NbrNInteg, out totInteg);
                NIntegPrices = Compteur(totInteg, prices, "sensonicRadio", true,false);
                TotInteg = totInteg;
            }

            if (!string.IsNullOrEmpty(NbrNEauCh) && !string.IsNullOrEmpty(NbrNEauFr))
            {
                int eauCh = 0;
                int.TryParse(NbrNEauCh, out eauCh);
                int eauFr = 0;
                int.TryParse(NbrNEauFr, out eauFr);

                totalEau = eauCh + eauFr;
                TotNEauch = eauCh;
                TotNEauFr = eauFr;
                totNEauch = eauCh;
                totNEauFr = eauFr;
                TotEau = totalEau;
            }
            else if (string.IsNullOrEmpty(NbrNEauCh) && !string.IsNullOrEmpty(NbrNEauFr))
            {

                int eauCh = 0;
                int.TryParse(NbrNEauCh, out eauCh);
                
                TotNEauch = eauCh;
                totalEau  = eauCh;
                TotEau = totalEau;
            }
            else if (!string.IsNullOrEmpty(NbrNEauCh) && string.IsNullOrEmpty(NbrNEauFr))
            {
                int eauFr = 0;
                int.TryParse(NbrNEauCh, out eauFr);

                TotNEauFr = eauFr;
                totalEau = eauFr;
                TotEau = totalEau;
            }

            NToEauPrices = Compteur(totalEau, prices, "domaquaRadio", false,false);

            if (!string.IsNullOrEmpty(NbrNEauCh))
            {
                int.TryParse(NbrNEauCh, out totNEauch);
                NEauChPrices = Compteur(totNEauch, prices, "domaquaRadio", false, false);
            }
            if (!string.IsNullOrEmpty(NbrNEauFr))
            {
                int.TryParse(NbrNEauFr, out totNEauFr);
                NEauFrPrices = Compteur(totNEauFr, prices, "domaquaRadio", false, false);
            }

            

            var doprimo3RadioChauf = NChaufPrices.Find(x => x.Id == "doprimoIIIRadio");
            var domaAquaEauTot = NToEauPrices.Find(x => x.Id == "domaquaRadio");
            
            var sensonic1_2Integ = NIntegPrices.Find(x => x.Id == "sensonicRadio1_2");
            var sensonic3_4Integ = NIntegPrices.Find(x => x.Id == "sensonicRadio3_4");

            string zeroVal = "N/A";

            int totalDevice = totNchauf + totalEau + totInteg;
            TotalDevice = totalDevice;

            float totalRadioRelevePrice = 0;
            float totaldomaqRelevePrice = 0;
            float totalSen12RelevePrice = 0;
            float totalSen34RelevePrice = 0;

            if (totNchauf == 0)
            {
                Dop3RadioChaufVente = zeroVal;
                Dop3RadioChaufRelev = zeroVal;
                Dop3RadioChaufLocat = zeroVal;

                Dop3RadioChaufVenteValue = 0;
                Dop3RadioChaufRelevValue = 0;
                Dop3RadioChaufLocatValue = 0;
            }
            else
            {
                Dop3RadioChaufVente = doprimo3RadioChauf.Vente.ToString();
                Dop3RadioChaufVenteValue = doprimo3RadioChauf.Vente;
                Dop3RadioChaufRelev = doprimo3RadioChauf.Releve.ToString();
                Dop3RadioChaufRelevValue = doprimo3RadioChauf.Releve;

                totalRadioRelevePrice = TotNchauf*Dop3RadioChaufRelevValue;

                if (totalDevice < 8)
                {
                    Dop3RadioChaufLocat = zeroVal;
                    Dop3RadioChaufLocatValue = 0;
                }
                else
                {
                    Dop3RadioChaufLocat = doprimo3RadioChauf.LocationMax15.ToString();
                    Dop3RadioChaufLocatValue = doprimo3RadioChauf.LocationMax15;
                }
                  
            }

            DomaquaTotal = totalEau.ToString();
            if (totalEau == 0)
            {
                DomaquaTotalReleveValue = 0;
                DomaquaTotalVenteValue = 0;
                DomaquaTotalLocatiValue = 0;

                DomaAquaEauChRVente = zeroVal;
                DomaAquaEauChRRelev = zeroVal;
                DomaAquaEauChRLocat = zeroVal;

                DomaAquaEauFrRVente = zeroVal;
                DomaAquaEauFrRRelev = zeroVal;
                DomaAquaEauFrRLocat = zeroVal;

                DomaquaTotalReleve = zeroVal;
                DomaquaTotalVente = zeroVal;
                DomaquaTotalLocati = zeroVal;
            }
            else
            {
                DomaquaTotalReleve = domaAquaEauTot.Releve.ToString();
                DomaquaTotalVente = domaAquaEauTot.Vente.ToString();

                DomaquaTotalReleveValue = domaAquaEauTot.Releve;
                DomaquaTotalVenteValue = domaAquaEauTot.Vente;

                totaldomaqRelevePrice = totalEau*DomaquaTotalReleveValue;

                if (TotalDevice< 4)
                {
                    DomaquaTotalLocati = zeroVal;
                    DomaquaTotalLocatiValue = 0;
                }
                else
                {
                    DomaquaTotalLocati = domaAquaEauTot.LocationMax15.ToString();
                    DomaquaTotalLocatiValue = domaAquaEauTot.LocationMax15;
                }

                if (totNEauFr == 0)
                {
                    DomaAquaEauFrRVente = zeroVal;
                    DomaAquaEauFrRRelev = zeroVal;
                    DomaAquaEauFrRLocat = zeroVal;
                }
                else
                {
                    DomaAquaEauFrRVente = DomaquaTotalVente;
                    DomaAquaEauFrRRelev = DomaquaTotalReleve;
                    DomaAquaEauFrRLocat = DomaquaTotalLocati;
                }

                if (totNEauch == 0)
                {
                    DomaAquaEauChRVente = zeroVal;
                    DomaAquaEauChRRelev = zeroVal;
                    DomaAquaEauChRLocat = zeroVal;
                }
                else
                {
                    DomaAquaEauChRVente = DomaquaTotalVente;
                    DomaAquaEauChRRelev = DomaquaTotalReleve;
                    DomaAquaEauChRLocat = DomaquaTotalLocati;
                }
            }

            if (totInteg == 0)
            {
                SensonicIn1_2Vente = zeroVal;
                SensonicIn1_2Relev = zeroVal;
                SensonicIn1_2Locat = zeroVal;

                SensonicIn3_4Vente = zeroVal;
                SensonicIn3_4Relev = zeroVal;
                SensonicIn3_4Locat = zeroVal;

                SensonicIn1_2RelevValue = 0;
                SensonicIn1_2LocatValue = 0;
                SensonicIn1_2VenteValue = 0;
                SensonicIn3_4VenteValue = 0;
                SensonicIn3_4RelevValue = 0;
                SensonicIn3_4LocatValue = 0;
            }
            else
            {
                SensonicIn1_2Vente = sensonic1_2Integ.Vente.ToString();
                SensonicIn1_2Relev = sensonic1_2Integ.Releve.ToString();

                SensonicIn3_4Vente = sensonic3_4Integ.Vente.ToString();
                SensonicIn3_4Relev = sensonic3_4Integ.Releve.ToString();

                SensonicIn1_2RelevValue = sensonic1_2Integ.Releve;
                SensonicIn1_2VenteValue = sensonic1_2Integ.Vente;
                SensonicIn3_4VenteValue = sensonic3_4Integ.Vente;
                SensonicIn3_4RelevValue = sensonic3_4Integ.Releve;

                totalSen12RelevePrice = totInteg*SensonicIn1_2RelevValue;
                totalSen34RelevePrice = totInteg*SensonicIn3_4RelevValue;

                if (TotalDevice < 2)
                {
                    SensonicIn1_2Locat = zeroVal;
                    SensonicIn3_4Locat = zeroVal;

                    SensonicIn1_2LocatValue = 0;
                    SensonicIn3_4LocatValue = 0;
                }
                else
                {
                    SensonicIn1_2Locat = sensonic1_2Integ.LocationMax15.ToString();
                    SensonicIn3_4Locat = sensonic3_4Integ.LocationMax15.ToString();

                    SensonicIn1_2LocatValue = sensonic1_2Integ.LocationMax15;
                    SensonicIn3_4LocatValue = sensonic3_4Integ.LocationMax15;
                }
            }

            float totalReleve = totalRadioRelevePrice + totaldomaqRelevePrice + totalSen12RelevePrice + totalSen34RelevePrice;
            if (totalReleve < 98)
            {
                if (totNchauf > 0)
                {
                    Dop3RadioChaufRelev = "97,75€ (Forfait)";
                    Dop3RadioChaufRelevValue = 97.75f;
                }
                if (totalEau > 0)
                    DomaquaTotalReleve = "97,75€ (Forfait)";
                if (totNEauch > 0)
                    DomaAquaEauChRRelev = "97,75€ (Forfait)";
                if (totNEauFr > 0)
                    DomaAquaEauFrRRelev = "97,75€ (Forfait)";
                if (totInteg > 0)
                {
                    SensonicIn1_2Relev = "97,75€ (Forfait)";
                    SensonicIn1_2RelevValue = 97.75f;
                }
                if (totInteg > 0)
                {
                    SensonicIn3_4Relev = "97,75€ (Forfait)";
                    SensonicIn3_4RelevValue = 97.75f;
                }
            }
        }
        #endregion

        #region Private Methods
        
        /// <summary>
        /// This method return a specific list of product price.
        /// </summary>
        /// <param name="nbrCompteur">Number of comptor.</param>
        /// <param name="prices">General list of price.</param>
        /// <param name="productName">Name of the product Analyzed.</param>
        /// <param name="sensonic">Check whether the porduct is sensonic.</param>
        /// <returns>List of tarifs for a specific product.</returns>
        private List<Tarifs> Compteur(int nbrCompteur, List<Tarifs> prices, string productName, bool sensonic, bool doprimo)
        {
            List<Tarifs> totalfinal = new List<Tarifs>();
            if (productName == ProductsIsta.doprimoIIIRadio.ToString() | productName == ProductsIsta.doprimoIIIRs.ToString())
            {
                if (1 < nbrCompteur && nbrCompteur < 8)
                {
                    totalfinal = TarifSelection(prices, 1);
                }
                else if (7 < nbrCompteur && nbrCompteur < 12)
                {
                    totalfinal = TarifSelection(prices, 8);
                }
                else if (11 < nbrCompteur && nbrCompteur < 16)
                {
                    totalfinal = TarifSelection(prices, 12);
                }
                else if (15 < nbrCompteur && nbrCompteur < 51)
                {
                    totalfinal = TarifSelection(prices, 16);
                }
                else if ( 50 < nbrCompteur && nbrCompteur < 251)
                {
                    totalfinal = TarifSelection(prices, 51);
                }
                else if (250 < nbrCompteur)
                {
                    totalfinal = TarifSelection(prices, 251);
                }
            }
            else if (productName == ProductsIsta.domaquaRadio.ToString())
            {
                if (0 < nbrCompteur && nbrCompteur < 4)
                {
                    totalfinal = TarifSelection(prices, 1);
                }
                else if (3 < nbrCompteur && nbrCompteur < 16)
                {
                    totalfinal = TarifSelection(prices, 4);
                }
                else if (15 < nbrCompteur && nbrCompteur < 150)
                {
                    totalfinal = TarifSelection(prices, 16);
                }
                else if (151 < nbrCompteur)
                {
                    totalfinal = TarifSelection(prices, 151);
                }
            }
            else if (productName == ProductsIsta.sensonicRadio1_2.ToString() | productName == ProductsIsta.sensonicRadio3_4.ToString() | productName == "sensonicRadio")
            {
                if (1 == nbrCompteur )
                {
                    totalfinal = TarifSelection(prices, 1);
                }
                if (1 < nbrCompteur && nbrCompteur < 16)
                {
                    totalfinal = TarifSelection(prices, 2);
                }
                else if (15 < nbrCompteur && nbrCompteur < 150)
                {
                    totalfinal = TarifSelection(prices, 16);
                }
                else if (151 < nbrCompteur)
                {
                    totalfinal = TarifSelection(prices, 151);
                }
            }
            return totalfinal;            
        }
        /// <summary>
        /// This method return a list of product price.
        /// </summary>
        /// <param name="prices">List of the prices of the products.</param>
        /// <param name="quantity">Quantity of the products</param>
        /// <returns>Return the prices of a typical product.</returns>
        private List<Tarifs> TarifSelection(IEnumerable<Tarifs> prices, int quantity)
        {
            var values = prices.Where(c => c.Value == quantity);
            return values.ToList();
        }
        #endregion
    }
}
