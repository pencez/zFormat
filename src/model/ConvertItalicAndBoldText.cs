using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//using System.Text.RegularExpressions;
//using System.Xml.Linq;
//using System.Xml;

namespace zFormat.model
{
    class ConvertItalicAndBoldText
    {

        static string lastRun = "1";
        static string currRun = "";

        public static void ConvertProc(FileInfo fileName)
        {
            using (var doc = WordprocessingDocument.Open(fileName.FullName, true))
            {
                foreach (var paragraph in doc.MainDocumentPart.RootElement.Descendants<Paragraph>())
                {
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        currRun = "";   //reset after each run
                        
                        if (run.RunProperties != null && run.RunProperties.Bold != null && 
                            (run.RunProperties.Bold.Val == null || run.RunProperties.Bold.Val))
                        {
                            currRun = "bold";
                        }
                        else if (run.RunProperties != null && run.RunProperties.Italic != null && 
                            (run.RunProperties.Italic.Val == null || run.RunProperties.Italic.Val))
                        {
                            currRun = "italics";
                        }
                        else if (run.RunProperties != null && run.RunProperties.Underline != null)
                        {
                            currRun = "underline";
                        }

                        if (lastRun != "" || lastRun != "1")
                        {
                            RunMarkup(run, "end", currRun);
                        }

                        if (currRun != "")
                        {
                            RunMarkup(run, "start", currRun);
                        }

                        
                    }
                }
            }
        }
        static void RunMarkup(Run run, string position, string action)
        {
            string text = run.Elements<Text>().Aggregate("", (s, t) => s + t.Text);
            if (action == "bold")
            {
                if (position == "start")
                {
                    run.PrependChild(new Text("BBB^"));
                    lastRun = "bold";
                }
                else
                {
                    run.AppendChild(new Text("&BBB"));
                    lastRun = "";
                }
            }
            else if (action == "italics")
            {
                if (position == "start")
                {
                    run.PrependChild(new Text("QQQ^"));
                    lastRun = "italics";
                }
                else
                {
                    run.AppendChild(new Text("&QQQ"));
                    lastRun = "";
                }
            }
            else if (action == "underline")
            {
                if (position == "start")
                {
                    run.PrependChild(new Text("UUU^"));
                    lastRun = "underline";
                }
                else
                {
                    run.AppendChild(new Text("&UUU"));
                    lastRun = "";
                }
            }
        }
    }
}
