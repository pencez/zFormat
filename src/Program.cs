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
        static string testDocx = @"C:\VS_Projects\Test\Test2.docx";
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

            //run bold, italics and underlines function
            //zFormat.model.ConvertItalicAndBoldText.ConvertProc(newDoc);
            //zFormat.model.FontMaster.SetRunFont(newDoc);
            //zFormat.model.StylesMaster.ExtractStylesPart(newDoc.FullName, false);

            zFormat.model.DocProcessor.processDoc(newDoc);
            zFormat.model.HeadersFooters.beginProc(newDoc);

            Console.WriteLine("*** All Done! ***");
            Console.ReadLine();

        }
    }
}
