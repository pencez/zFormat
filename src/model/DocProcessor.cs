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

namespace zFormat.model
{
    class DocProcessor
    {
        // Process the actual document, one paragraph at a time
        public static void processDoc(FileInfo newDoc)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(newDoc.FullName, true))
            {

                var docBody = doc.MainDocumentPart.Document.Body;
                var paragraphs = zFormat.model.GetMetrics.paraCount;
                //var pChapter = zFormat.model.SearchAndReplace.chapElement.Distinct().ToList();


                int c = 1;
                for (var i = 0; i < paragraphs; i++)
                {
                    if (docBody.Descendants<Paragraph>().ElementAtOrDefault(i) != null)
                    {
                        Paragraph p = docBody.Descendants<Paragraph>().ElementAt(i);
                        var eleText = docBody.Descendants<Paragraph>().ElementAtOrDefault(i).InnerText;
                        if (eleText == "Chapter " + c)
                        {
                            // Paragraph number for Chapter headings
                            zFormat.model.StylesMaster.ApplyStyleToParagraph(doc, "zHeading", "zHeading", p);
                            // Call function to check for page break before chapter heading
                            zFormat.model.PageControls.checkForPageBreak(doc, p, i);
                            // At start of chapter, check for no ind -- fix if needed
                            //zFormat.model.PageControls.checkForIndent(doc, i, "N");
                            c++;
                        }
                        else
                        {
                            // Set appropriate style to docBody
                            zFormat.model.StylesMaster.ApplyStyleToParagraph(doc, "zNormal", "zNormal", p);
                            // At start of chapter, check for no ind -- fix if needed
                            zFormat.model.PageControls.checkForIndent(doc, i, "Y");
                        }
                    }
                }
                //paragraphList.Distinct().ToList().ForEach(Console.WriteLine);


            }
        }


        

    }
}
