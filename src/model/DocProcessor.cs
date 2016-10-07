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
        // Process the document, one paragraph at a time
        public static void processDoc(FileInfo newDoc)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(newDoc.FullName, true))
            {

                List<int> paragraphList = new List<int>();
                var paragraphs = zFormat.model.SearchAndReplace.paraCount;
                var pChapter = zFormat.model.SearchAndReplace.chapElement.Distinct().ToList();

                int c = 0;
                for (var i = 0; i < paragraphs; i++)
                {
                    if (pChapter.Contains(i))
                    {
                        // Paragraph number for Chapter headings
                        Paragraph p = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>().ElementAtOrDefault(pChapter[c]);
                        zFormat.model.StylesMaster.ApplyStyleToParagraph(doc, "zHeading", "zHeading", p);
                        // Call function to check for page break before chapter heading
                        checkForPageBreak(doc, p, pChapter[c]);
                        c++;
                    }
                    else
                    {
                        if (doc.MainDocumentPart.Document.Body.Descendants<Paragraph>().ElementAtOrDefault(i) != null)
                        {
                            Paragraph p = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>().ElementAt(i);
                            zFormat.model.StylesMaster.ApplyStyleToParagraph(doc, "zNormal", "zNormal", p);
                        }
                    }
                }
                //paragraphList.Distinct().ToList().ForEach(Console.WriteLine);


            }
        }


        public static void checkForPageBreak(WordprocessingDocument wDoc, Paragraph para, int e)
        {
            var runs = new List<Run>();            
                //runs = para.OfType<Run>()
                //    .Where(W.lastRenderedPageBreak).ToList();
                //    r.RunProperties.RunStyle.Val.Value.Contains("lastRenderedPageBreak") ||      
                //var zTest = runs[0].LocalName;


            var zTemp = para.OuterXml;
            var wordText = para.InnerText;
            var pageBreak = Regex.Match(zTemp, "w:lastRenderedPageBreak");

                



            var xDoc = wDoc.MainDocumentPart.GetXDocument();
            IEnumerable<XElement> content;

            // Match content from prargraphProperties
            content = xDoc.Descendants(W.p);
            //var pageBreak = content.Elements(W.r).Elements(W.lastRenderedPageBreak);

            
        }


    }
}
