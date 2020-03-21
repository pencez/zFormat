using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace zFormat
{
    class Program
    {
        static string testDocx = @"C:\Temp\ebook\oneoclock_small_zTest.docx";
        //static string testDocx = @"C:\Temp\ebook\GOSH_zTest.docx";
        //static string testDocx = @"C:\VS_Projects\Test\BigTest.docx";

        static void Main(string[] args)
        {
            
            var n = DateTime.Now;
            //var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput"));
            bool exists = System.IO.Directory.Exists("ExampleOutput");
            if (!exists) { tempDi.Create(); }

            var sourceDoc = new FileInfo(testDocx);
            var newDoc = new FileInfo(Path.Combine(tempDi.FullName, "z_" + DateTime.Now.Ticks + "_" + sourceDoc.Name));
            File.Copy(sourceDoc.FullName, newDoc.FullName);

            
            //run metrics report function
            zFormat.model.GetMetrics.contentVitals(newDoc);
            //get the fonts being used
            zFormat.model.FontMaster.GetFontList(newDoc);
            //set the font for doc
            zFormat.model.FontMaster.SetRunFont(newDoc, "Georgia");
            //set the font for chapter headings
            zFormat.model.FontMaster.SetHeadFont(newDoc, "Rockwell Extra Bold");



            //Replace current font with Georgia font
            //DocumentFormat.OpenXml.Packaging.WordprocessingDocument wDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(newDoc.FullName, true);
            //zFormat.model.SearchAndReplacer.SearchAndReplace(wDoc, "Verdana", "Georgia", true);


            //run bold, italics and underlines function
            //zFormat.model.ConvertItalicAndBoldText.ConvertProc(newDoc);
            //zFormat.model.FontMaster.SetRunFont(newDoc);
            //zFormat.model.StylesMaster.ExtractStylesPart(newDoc.FullName, false);

            //zFormat.model.DocProcessor.processDoc(newDoc);
            //zFormat.model.HeadersFooters.beginProc(newDoc);

            Console.WriteLine("*** All Done! ***");
            Console.ReadLine();

        }
    }
}
