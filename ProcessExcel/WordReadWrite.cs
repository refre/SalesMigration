using System;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace Ista.Migration.Excel
{
    /// <summary>
    /// http://www.techrepublic.com/blog/howdoi/how-do-i-modify-word-documents-using-c/190
    /// </summary>
    public class WordReadWrite
    {
        /// <summary>
        /// Method to create the template word and replace all tags by values.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="saveAs"></param>
        /// <param name="currentElement"></param>
        public void CreatWordDocument(object fileName, object saveAs, MigrationElement currentElement)
        {
            try
            {
                object missing = System.Reflection.Missing.Value;

                Application wordApp = new Application();

                //next line gives a bug : 
                // http://stackoverflow.com/questions/2483659/interop-type-cannot-be-embedded
                //Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();

                Microsoft.Office.Interop.Word.Document aDoc = null;

                if (File.Exists((string)fileName))
                {
                    DateTime today = DateTime.Now;
                    object readOnly = false;
                    object isVisible = false;

                    wordApp.Visible = false;

                    aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing,
                                           ref missing, ref missing, ref missing, ref missing, ref missing,
                                           ref missing, ref isVisible, ref missing, ref missing, ref missing,
                                           ref missing);
                    aDoc.Activate();

                    FindAndReplace(wordApp, "CodeNameGerant",currentElement.Element.CodeNameGerant);
                    FindAndReplace(wordApp, "CodeAdresseGerant", currentElement.Element.CodeAdresseGerant);
                    FindAndReplace(wordApp, "CodePostauxGerant", currentElement.CodePostauxGerant);
                    FindAndReplace(wordApp, "LocaliteGerant", currentElement.LocaliteGerant);
                    FindAndReplace(wordApp, "CodeOffreNum", currentElement.Element.CodeOffreNum);
                    FindAndReplace(wordApp, "CodeImmeuble", currentElement.Element.CodeImmeuble);
                    FindAndReplace(wordApp, "CodeAdresseImeuble", currentElement.Element.TotalAdressImeuble);
                   
                    if (currentElement.Element.TotNchauf > 0)
                    {
                        /*******************/
                        /** Location part **/
                        /*******************/
                        FindAndReplace(wordApp, "CodeDPRLoc", currentElement.ChauffageNombreNRad);
                        FindAndReplace(wordApp, "PrixDPRLoc", currentElement.Element.Dop3RadioChaufLocat + " €");
                        /*************/
                        /** Selling **/
                        /*************/
                        FindAndReplace(wordApp, "CodeDPRVente", currentElement.ChauffageNombreNRad);
                        FindAndReplace(wordApp, "PrixDPRVente", currentElement.Element.Dop3RadioChaufVente + " €");
                        /*************/
                        /** Reading **/
                        /*************/
                        FindAndReplace(wordApp, "CodeRep", currentElement.ChauffageNombreNRad);
                        FindAndReplace(wordApp, "PrixRep", currentElement.Element.Dop3RadioChaufRelev == "97,75€ (Forfait)" ? currentElement.Element.Dop3RadioChaufRelev : currentElement.Element.Dop3RadioChaufRelev + " €");
                    }
                    else
                    {
                        /*******************/
                        /** Location part **/
                        /*******************/
                        FindAndReplace(wordApp, "CodeDPRLoc", "");
                        FindAndReplace(wordApp, "PrixDPRLoc", "");
                        /*************/
                        /** Selling **/
                        /*************/
                        FindAndReplace(wordApp, "CodeDPRVente", "");
                        FindAndReplace(wordApp, "PrixDPRVente", "");
                        /*************/
                        /** Reading **/
                        /*************/
                        FindAndReplace(wordApp, "CodeRep", "");
                        FindAndReplace(wordApp, "PrixRep", "");                      
                    }
                    /*******************/
                    /** Location part **/
                    /*******************/
                    FindAndReplace(wordApp, "CodeDPRSLoc", "");
                    FindAndReplace(wordApp, "PrixDPRSLoc", "");
                    /*************/
                    /** Selling **/
                    /*************/
                    FindAndReplace(wordApp, "CodeDPRSVente", "");
                    FindAndReplace(wordApp, "PrixDPRSVente", "");
                                       
                    /*******************/
                    /** Location part **/
                    /*******************/
                    FindAndReplace(wordApp, "CodeDomaquaLoc", "");
                    FindAndReplace(wordApp, "PrixDomaquaLoc", "");
                    FindAndReplace(wordApp, "CodeDomaquaChLoc", "");
                    FindAndReplace(wordApp, "PrixDomaquaChLoc", "");
                    /*************/
                    /** Selling **/
                    /*************/
                    FindAndReplace(wordApp, "CodeDomaquaVente", "");
                    FindAndReplace(wordApp, "PrixDomaquaVente", "");
                    FindAndReplace(wordApp, "CodeDomaquaChVente", "");
                    FindAndReplace(wordApp, "PrixDomaquaChVente", "");

                    if (currentElement.Element.TotNEauFr > 0)
                    {
                        /*******************/
                        /** Location part **/
                        /*******************/

                        //06/03/2013: changement demandé par edouard!! si la location n'est pas permise (c'est à dire en dessous de la limite)
                        //Alors les valeurs ne doivent pas se voir.

                        if (currentElement.Element.DomaAquaEauFrRLocat == "0")
                        {
                            FindAndReplace(wordApp, "CodeDomaquaRLoc", "");
                            FindAndReplace(wordApp, "PrixDomaquaRLoc", "");
                        }
                        else
                        {
                            FindAndReplace(wordApp, "CodeDomaquaRLoc", currentElement.EauFroideNombreNRad);
                            FindAndReplace(wordApp, "PrixDomaquaRLoc", currentElement.Element.DomaAquaEauFrRLocat + " €");
                        }

                        
                        /*************/
                        /** Selling **/
                        /*************/
                        FindAndReplace(wordApp, "CodeDomaquaRVente", currentElement.EauFroideNombreNRad);
                        FindAndReplace(wordApp, "PrixDomaquaRVente", currentElement.Element.DomaAquaEauFrRVente + " €");
                    }
                    else
                    {
                        /*******************/
                        /** Location part **/
                        /*******************/
                        FindAndReplace(wordApp, "CodeDomaquaRLoc", "");
                        FindAndReplace(wordApp, "PrixDomaquaRLoc", "");
                        /*************/
                        /** Selling **/
                        /*************/
                        FindAndReplace(wordApp, "CodeDomaquaRVente","");
                        FindAndReplace(wordApp, "PrixDomaquaRVente","");
                    }
                    /*******************/
                    /** Location part **/
                    /*******************/
                    FindAndReplace(wordApp, "CodeDomaquaChLoc", "");
                    FindAndReplace(wordApp, "PrixDomaquaLoc", "");
                    /*************/
                    /** Selling **/
                    /*************/
                    FindAndReplace(wordApp, "CodeDomaquaChVente", "");
                    FindAndReplace(wordApp, "PrixDomaquaVente", "");
                    if (currentElement.Element.TotNEauch > 0)
                    {
                        /*******************/
                        /** Location part **/
                        /*******************/

                        //06/03/2013: changement demandé par edouard!! si la location n'est pas permise (c'est à dire en dessous de la limite)
                        //Alors les valeurs ne doivent pas se voir.
                        if (currentElement.Element.DomaAquaEauChRLocat=="0")
                        {
                            FindAndReplace(wordApp, "CodeDomaquaChRLoc","");
                            FindAndReplace(wordApp, "PrixDomaquaChRLoc", ""); 
                        }
                        else
                        {
                            FindAndReplace(wordApp, "CodeDomaquaChRLoc", currentElement.EauChaudeNombreNRad);
                            FindAndReplace(wordApp, "PrixDomaquaChRLoc", currentElement.Element.DomaAquaEauChRLocat + " €");
                        }
                        /*************/
                        /** Selling **/
                        /*************/
                        FindAndReplace(wordApp, "CodeDomaquaChRVente", currentElement.EauChaudeNombreNRad);
                        FindAndReplace(wordApp, "PrixDomaquaChRVente", currentElement.Element.DomaAquaEauChRVente + " €");
                    }
                    else
                    {
                        /*******************/
                        /** Location part **/
                        /*******************/
                        FindAndReplace(wordApp, "CodeDomaquaChRLoc", "");
                        FindAndReplace(wordApp, "PrixDomaquaChRLoc", "");
                        /*************/
                        /** Selling **/
                        /*************/
                        FindAndReplace(wordApp, "CodeDomaquaChRVente", "");
                        FindAndReplace(wordApp, "PrixDomaquaChRVente", "");
                    }                   
                    if (currentElement.Element.TotInteg > 0)
                    {
                        /*******************/
                        /** Location part **/
                        /*******************/
                        FindAndReplace(wordApp, "CodeSensonicLoc", currentElement.IntegrateurNombreNRad);
                        FindAndReplace(wordApp, "PrixSensonicLoc", currentElement.Element.SensonicIn1_2Locat + " €");
                        FindAndReplace(wordApp, "CodeSensonicRLoc", currentElement.IntegrateurNombreNRad);
                        FindAndReplace(wordApp, "PrixSensonicRLoc", currentElement.Element.SensonicIn3_4Locat + " €");
                        /*************/
                        /** Selling **/
                        /*************/
                        FindAndReplace(wordApp, "CodeSensonicVente", currentElement.IntegrateurNombreNRad);
                        FindAndReplace(wordApp, "PrixSensonicVente", currentElement.Element.SensonicIn1_2Vente + " €");
                        FindAndReplace(wordApp, "CodeSensonicRVente", currentElement.IntegrateurNombreNRad);
                        FindAndReplace(wordApp, "PrixSensonicRVente", currentElement.Element.SensonicIn3_4Vente + " €");
                    }
                    else
                    {
                        /*******************/
                        /** Location part **/
                        /*******************/
                        FindAndReplace(wordApp, "CodeSensonicLoc", "");
                        FindAndReplace(wordApp, "PrixSensonicLoc", "");
                        FindAndReplace(wordApp, "CodeSensonicRLoc", "");
                        FindAndReplace(wordApp, "PrixSensonicRLoc", "");
                        /*************/
                        /** Selling **/
                        /*************/
                        FindAndReplace(wordApp, "CodeSensonicVente", "");
                        FindAndReplace(wordApp, "PrixSensonicVente", "");
                        FindAndReplace(wordApp, "CodeSensonicRVente", "");
                        FindAndReplace(wordApp, "PrixSensonicRVente", "");
                    }
                    if (currentElement.Element.TotEau > 0)
                    {
                        FindAndReplace(wordApp, "CodeEau", currentElement.Element.TotEau.ToString());
                        FindAndReplace(wordApp, "PrixEau", currentElement.Element.DomaquaTotalReleve == "97,75€ (Forfait)" ? currentElement.Element.DomaquaTotalReleve : currentElement.Element.DomaquaTotalReleve + " €");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "CodeEau", "");
                        FindAndReplace(wordApp, "PrixEau", "");
                    }
                    if (currentElement.Element.TotInteg > 0)
                    {
                        FindAndReplace(wordApp, "CodeIntegr", currentElement.IntegrateurNombreNRad);

                        FindAndReplace(wordApp, "PrixIntegr", currentElement.Element.SensonicIn1_2Relev == "97,75€ (Forfait)" ? currentElement.Element.SensonicIn1_2Relev : currentElement.Element.SensonicIn1_2Relev + " €");

                        FindAndReplace(wordApp, "CodeElec", currentElement.IntegrateurNombreNRad);
                        FindAndReplace(wordApp, "PrixElec", currentElement.Element.SensonicIn3_4Relev == "97,75€ (Forfait)" ? currentElement.Element.SensonicIn3_4Relev : currentElement.Element.SensonicIn3_4Relev + " €");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "CodeIntegr", "");
                        FindAndReplace(wordApp, "PrixIntegr", "");
                        FindAndReplace(wordApp, "CodeElec", "");
                        FindAndReplace(wordApp, "PrixElec", "");
                    }
                }
                else
                {
                    Console.WriteLine("Error");
                    return;
                }

                aDoc.SaveAs(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                aDoc.Close(ref missing, ref missing, ref missing);

                Console.WriteLine("File Created.");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Method to find a tag and replace by a value.
        /// </summary>
        /// <param name="wordApp"></param>
        /// <param name="findText"></param>
        /// <param name="replaceWithText"></param>
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object macthAllWordFroms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards,
                                          ref matchSoundsLike,
                                          ref matchWholeWord, ref forward, ref wrap, ref format, ref replaceWithText,
                                          ref replace, ref matchKashida, ref matchDiacritics,
                                          ref matchAlefHamza, ref matchControl);
        }
    }
}