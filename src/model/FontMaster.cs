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
using System.Windows;

namespace zFormat.model
{
    class FontMaster
    {

        // Set the font for a text run.
        public static void GetFontList(FileInfo fileName)
        {
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fileName.FullName, true))
            {
                var fontList = wDoc.MainDocumentPart.FontTablePart.Fonts.Elements<Font>();
                //.Select(
                //       Function(c) If(c.Ascii.HasValue, c.Ascii.InnerText, String.Empty)).Distinct().ToList()

                //fontList.AddRange(runFonts)
                String theFonts = "";
                foreach (var zfont in fontList)
                {
                    theFonts = theFonts + zfont.Name + ",";
                }
                theFonts = theFonts.TrimEnd(',');
                Console.WriteLine("Fonts Used: {0}", theFonts);

            }
        }


        // Set the font for a text headings.
        public static void SetHeadFont(FileInfo fileName, String hFont)
        {
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fileName.FullName, true))
            {
                Body body = wDoc.MainDocumentPart.Document.Body;
                //Get all paragraphs
                var lstParagrahps = body.Descendants<Paragraph>().ToList();
                foreach (var para in lstParagrahps)
                {
                    var subRuns = para.Descendants<Run>().ToList();
                    foreach (var run in subRuns)
                    {
                        if (run.InnerText.StartsWith("Chapter ") || run.InnerText.StartsWith("CHAPTER "))
                        {
                            var subRunProp = run.Descendants<RunProperties>().ToList().FirstOrDefault();

                            var newFont = new RunFonts();
                            newFont.Ascii = hFont;
                            //newFont.EastAsia = hFont;
                            newFont.HighAnsi = hFont;
                            //newFont.ComplexScript = hFont;

                            if (subRunProp != null)
                            {
                                var font = subRunProp.Descendants<RunFonts>().FirstOrDefault();
                                subRunProp.ReplaceChild<RunFonts>(newFont, font);
                            }
                            else
                            {
                                var tmpSubRunProp = new RunProperties();
                                tmpSubRunProp.AppendChild<RunFonts>(newFont);
                                run.AppendChild<RunProperties>(tmpSubRunProp);
                            }

                        }
                    }
                }
                wDoc.MainDocumentPart.Document.Save();
                wDoc.Close();
            }
        }

        // Set the font for a text run.
        public static void SetRunFont(FileInfo fileName, String zFont)
        {
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fileName.FullName, true))
            {
                Body body = wDoc.MainDocumentPart.Document.Body;
                //Get all paragraphs
                var lstParagrahps = body.Descendants<Paragraph>().ToList();
                foreach (var para in lstParagrahps)
                {
                    var subRuns = para.Descendants<Run>().ToList();
                    foreach (var run in subRuns)
                    {
                        var subRunProp = run.Descendants<RunProperties>().ToList().FirstOrDefault();

                        var newFont = new RunFonts();
                        newFont.Ascii = zFont;
                        newFont.EastAsia = zFont;
                        newFont.HighAnsi = zFont;
                        newFont.ComplexScript = zFont;

                        if (subRunProp != null)
                        {
                            var font = subRunProp.Descendants<RunFonts>().FirstOrDefault();
                            subRunProp.ReplaceChild<RunFonts>(newFont, font);
                        }
                        else
                        {
                            var tmpSubRunProp = new RunProperties();
                            tmpSubRunProp.AppendChild<RunFonts>(newFont);
                            run.AppendChild<RunProperties>(tmpSubRunProp);
                        }

                    }
                }
                wDoc.MainDocumentPart.Document.Save();
                wDoc.Close();
            }
        }
    }
}
