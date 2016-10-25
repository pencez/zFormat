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

namespace zFormat.model.headers
{
    class SectionsMaster
    {
        public static void RemoveSectionBreaks(string filename)
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                List<ParagraphProperties> paraProps = mainPart.Document.Descendants<ParagraphProperties>()
                .Where(pPr => IsSectionProps(pPr)).ToList();

                foreach (ParagraphProperties pPr in paraProps)                {
                    pPr.RemoveChild<SectionProperties>(pPr.GetFirstChild<SectionProperties>());
                }
                mainPart.Document.Save();
            }
        }


        public static void CreateSectionBreaks(string filename)
        {

        }


        static bool IsSectionProps(ParagraphProperties pPr)
        {
            SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();
            if (sectPr == null)
                return false;
            else
                return true;
        }
    }
}
