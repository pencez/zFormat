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
    class GetMetrics
    {
        public static int paraCount = 0;
        public static int chapCount = 0;
        public static int firstLine = 0;
        public static List<int> chapElement = new List<int>();


        // To search and track content vital stats
        public static void contentVitals(FileInfo newDoc)
        {

            /*
            var n = DateTime.Now;
            //var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput"));
            bool exists = System.IO.Directory.Exists("ExampleOutput");
            if (!exists) { tempDi.Create(); }

            var sourceDoc = new FileInfo(document);
            var newDoc = new FileInfo(Path.Combine(tempDi.FullName, "z_" + DateTime.Now.Ticks + "_" + sourceDoc.Name));
            File.Copy(sourceDoc.FullName, newDoc.FullName);
            */


            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(newDoc.FullName, true))
            {
                var xDoc = wDoc.MainDocumentPart.GetXDocument();
                Regex regex;
                IEnumerable<XElement> content;
                content = xDoc.Descendants(W.p);

                // Count number of pages
                //var pageCount = wDoc.ExtendedFilePropertiesPart.Properties.Pages.InnerText.ToString();
                //var pageCount = wDoc.ExtendedFilePropertiesPart.Properties.Pages.Text.ToString();

                // Count chapters
                chapCount = 1;
                firstLine = 1;
                foreach (var element in wDoc.MainDocumentPart.Document.Body) {
                    if (element.InnerXml.IndexOf("<w:br w:type=\"page\" />") != -1)
                    {
                        chapCount++;
                    }                   
                }


                // Count paragraphs
                regex = new Regex("[.]\x020+");
                paraCount = OpenXmlRegex.Match(content, regex);
                // Set paragraphs in doc
                //int i = 0;
                //foreach (var para in content)
                //{
                //    var newPara = (XElement)TransformEnvironmentNewLineToParagraph(para, i);
                //    para.ReplaceNodes(newPara.Nodes());
                //    i++;
                //}
                //wDoc.MainDocumentPart.PutXDocument();

                // Count underlines, bold and italics
                var underlines = content.Elements(W.r).Elements(W.rPr).Elements(W.u).Attributes(W.val);
                var boldness = content.Elements(W.r).Elements(W.rPr).Elements(W.b);
                var italics = content.Elements(W.r).Elements(W.rPr).Elements(W.i);
                var uCount = underlines.Count();
                var bCount = boldness.Count();
                var iCount = italics.Count();

                //Console.WriteLine("Page Count: " + pageCount);
                Console.WriteLine("Chapter Count: {0}", chapCount);
                Console.WriteLine("FirstLine Count: {0}", firstLine);
                Console.WriteLine("Paragraph Count: {0}", paraCount);
                Console.WriteLine("Underlines Count: {0}", uCount);
                Console.WriteLine("Boldness Count: {0}", bCount);
                Console.WriteLine("Italics Count: {0}", iCount);
                //chapElement.Distinct().ToList().ForEach(Console.WriteLine);
                //chapElement.ForEach(Console.WriteLine);

                wDoc.Close();
                // Call to get Style counts and names
                //zFormat.model.StylesMaster.getStylesInfo(newDoc);

            }      
        
        }

        private static object TransformEnvironmentNewLineToParagraph(XNode node, int ele)
        {
            var element = node as XElement;
            if (element != null)
            {
                
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformEnvironmentNewLineToParagraph(n, ele)));

            }
            return node;
        }

    }
}
