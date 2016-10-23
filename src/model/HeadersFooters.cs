using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;

namespace zFormat.model
{
    class HeadersFooters
    {
        
        public static string filepathFrom = @"C:\VS_Projects\Test\HF_Template1.docx";

        public static void beginProc(FileInfo newDoc)
        {
            // Get info for headers/footers
            //get book title, header/footer info from website -- use vars for now
            string bookTitle = "Something's Cooking";
            string bookAuthor = "Joanne Pence";
            // Header/Footer options
            /*  option1 = 
             *  option2 = 
             *  option3 = 
             */
            string hfOption = "option1";
            string pageNumLocation = "footer";

            processHeader(newDoc, bookTitle, bookAuthor, hfOption);
            if (pageNumLocation == "footer")
            {
                processFooter(newDoc);
            }

        }

        
            
        public static void processHeader(FileInfo newDoc, string title, string author, string hfOption)
        {

            // Replace header in target document with header of source document.
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(newDoc.FullName, true))
            {
                MainDocumentPart mainPart = wdDoc.MainDocumentPart;

                // Delete the existing header part.
                mainPart.DeleteParts(mainPart.HeaderParts);

                // Create a new header part.
                DocumentFormat.OpenXml.Packaging.HeaderPart headerPart1 = mainPart.AddNewPart<HeaderPart>();
                DocumentFormat.OpenXml.Packaging.HeaderPart headerPart2 = mainPart.AddNewPart<HeaderPart>();

                // Get Id of the headerPart.
                string rId1 = mainPart.GetIdOfPart(headerPart1);
                string rId2 = mainPart.GetIdOfPart(headerPart2);

                // Feed target headerPart with source headerPart.
                using (WordprocessingDocument wdDocSource = WordprocessingDocument.Open(filepathFrom, true))
                {
                    // Get first header and replace template author with actual author
                    DocumentFormat.OpenXml.Packaging.HeaderPart firstHeader = wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();
                    string headerTemplateAuthor = firstHeader.Header.InnerText;
                    firstHeader.Header.InnerXml = firstHeader.Header.InnerXml.Replace(headerTemplateAuthor, author);
                    wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();
                    // Get second header and replace template title with actual title
                    DocumentFormat.OpenXml.Packaging.HeaderPart secondHeader = wdDocSource.MainDocumentPart.HeaderParts.ElementAtOrDefault(1);
                    string headerTemplateTitle = secondHeader.Header.InnerText;
                    secondHeader.Header.InnerXml = secondHeader.Header.InnerXml.Replace(headerTemplateTitle, title);
                    wdDocSource.MainDocumentPart.HeaderParts.ElementAtOrDefault(1);

                    if (firstHeader != null)
                    {
                        headerPart1.FeedData(firstHeader.GetStream());
                    }
                    if (secondHeader != null)
                    {
                        headerPart2.FeedData(secondHeader.GetStream());
                    }
                    wdDocSource.Save();
                }

                // Get SectionProperties and Replace HeaderReference with new Id.
                IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs = mainPart.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    // Delete existing references to headers.
                    sectPr.RemoveAllChildren<HeaderReference>();

                    // Create the new header reference node.
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId1, Type = HeaderFooterValues.Default });
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId2, Type = HeaderFooterValues.Even });
                }

                // Call function to add the header settings to the settings.xml reference
                AddSettingsToMainDocumentPart(mainPart, "header");

            }          
            
        }



        public static void processFooter(FileInfo newDoc)
        {
            // Replace header in target document with header of source document.
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(newDoc.FullName, true))
            {
                MainDocumentPart mainPart = wdDoc.MainDocumentPart;

                // Delete the existing footer part.
                mainPart.DeleteParts(mainPart.FooterParts);

                // Create a new footer part.
                DocumentFormat.OpenXml.Packaging.FooterPart footerPart1 = mainPart.AddNewPart<FooterPart>();
                DocumentFormat.OpenXml.Packaging.FooterPart footerPart2 = mainPart.AddNewPart<FooterPart>();
                DocumentFormat.OpenXml.Packaging.FootnotesPart footnotesPart = mainPart.AddNewPart<FootnotesPart>();
                DocumentFormat.OpenXml.Packaging.EndnotesPart endnotesPart = mainPart.AddNewPart<EndnotesPart>();

                // Get Id of the headerPart.
                string rId1 = mainPart.GetIdOfPart(footerPart1);
                string rId2 = mainPart.GetIdOfPart(footerPart2);

                // Feed target headerPart with source headerPart.
                using (WordprocessingDocument wdDocSource = WordprocessingDocument.Open(filepathFrom, true))
                {
                    // Get first footer
                    DocumentFormat.OpenXml.Packaging.FooterPart firstFooter = wdDocSource.MainDocumentPart.FooterParts.FirstOrDefault();
                    wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();
                    // Get second footer
                    DocumentFormat.OpenXml.Packaging.FooterPart secondFooter = wdDocSource.MainDocumentPart.FooterParts.ElementAtOrDefault(1);
                    wdDocSource.MainDocumentPart.HeaderParts.ElementAtOrDefault(1);

                    if (firstFooter != null)
                    {
                        footerPart1.FeedData(firstFooter.GetStream());
                    }
                    if (secondFooter != null)
                    {
                        footerPart2.FeedData(secondFooter.GetStream());
                    }
                    //wdDocSource.Save();

                    // Now build Footnotes and Endnotes xml files
                    DocumentFormat.OpenXml.Packaging.FootnotesPart footerNotes = wdDocSource.MainDocumentPart.FootnotesPart;
                    DocumentFormat.OpenXml.Packaging.EndnotesPart endNotes = wdDocSource.MainDocumentPart.EndnotesPart;

                    if (footerNotes != null)
                    {
                        footnotesPart.FeedData(footerNotes.GetStream());
                    }
                    if (endNotes != null)
                    {
                        endnotesPart.FeedData(endNotes.GetStream());
                    }

                }

                // Get SectionProperties and Replace HeaderReference with new Id.
                IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs = mainPart.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    // Delete existing references to headers.
                    sectPr.RemoveAllChildren<FooterReference>();

                    // Create the new header reference node.
                    sectPr.PrependChild<FooterReference>(new FooterReference() { Id = rId1, Type = HeaderFooterValues.Default });
                    sectPr.PrependChild<FooterReference>(new FooterReference() { Id = rId2, Type = HeaderFooterValues.Even });
                }

                // Call function to add the header settings to the settings.xml reference
                AddSettingsToMainDocumentPart(mainPart, "footer");

            }
        }


        private static void AddSettingsToMainDocumentPart(MainDocumentPart part, string HeadFoot)
        {
            DocumentSettingsPart settingsPart = part.DocumentSettingsPart;
            if (settingsPart == null)
                settingsPart = part.AddNewPart<DocumentSettingsPart>();
            
            if (HeadFoot == "header"){        
                settingsPart.Settings.Append(            
                    new EvenAndOddHeaders(),
                    new HeaderShapeDefaults(
                        new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 2049 }
                    )
                );
            }
            else if (HeadFoot == "footer")
            {
                settingsPart.Settings.Append( 
                    new FootnoteProperties(
                        new Footnote() { Id = -1 },
                        new Footnote() { Id = 0 }
                    ),
                    new EndnoteProperties(
                        new Endnote() { Id = -1 },
                        new Endnote() { Id = 0 }
                    )
                );
            }
            settingsPart.Settings.Save();

        }

    }
}
