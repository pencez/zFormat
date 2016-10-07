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
    class FontMaster
    {
        // Set the font for a text run.
        public static void SetRunFont(FileInfo fileName)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Open(fileName.FullName, true))
            {
                // Use an object initializer for RunProperties and rPr.
                RunProperties rPr = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = "Arial"
                    });

                // Set the font to Arial to the first Run.
                Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
                r.PrependChild<RunProperties>(rPr);

                // Save changes to the MainDocumentPart part.
                package.MainDocumentPart.Document.Save();
            }
        }
    }
}
