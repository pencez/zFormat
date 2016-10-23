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
    class PageControls
    {
        public static void checkForPageBreak(WordprocessingDocument doc, Paragraph para, int e)
        {
            /* 
             *** Check for page breaks to add when appropriate
             ***  - add at start of next chapter, when missing
            */
            var wDoc = doc.MainDocumentPart.Document;
            var zTempOuterXml = para.OuterXml;
            var pbExistsTF = Regex.Match(zTempOuterXml, "w:lastRenderedPageBreak").Success;
            //var pbExistsTF = getPageBreak.Success;

            var lastRenderedPageBreak = "<w:lastRenderedPageBreak /><w:t>";
            var actualPageBreak = "</w:pPr>";

            //var wordText = para.InnerText;

            if (pbExistsTF == false)
            {
                //apply lastRenderedPageBreak to current run                
                var myCurrEle = wDoc.Body.Descendants<Paragraph>().ElementAt(e);
                myCurrEle.InnerXml = myCurrEle.InnerXml.Replace("<w:t>", lastRenderedPageBreak);
                //cleanup
                myCurrEle.InnerXml = myCurrEle.InnerXml.Replace(" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");

                if (e != 0)
                {
                    //apply actualPageBreak to previous run
                    var myPrevEle = wDoc.Body.Descendants<Paragraph>().ElementAt(e - 1);
                    var myPrevEleOuterXml = myPrevEle.OuterXml;

                    //check if a run exists already. May need to add line break first
                    var doesRunExistTF = Regex.Match(myPrevEleOuterXml, "<w:r>").Success;
                    if (doesRunExistTF == true)
                    {
                        actualPageBreak = "<w:br w:type=\"page\"/></w:r>";
                        myPrevEle.InnerXml = myPrevEle.InnerXml.Replace("</w:r>", actualPageBreak);
                    }
                    else
                    {
                        actualPageBreak = "</w:pPr><w:r><w:br w:type=\"page\"/></w:r>";
                        myPrevEle.InnerXml = myPrevEle.InnerXml.Replace("</w:pPr>", actualPageBreak);
                    }
                    //cleanup
                    myPrevEle.InnerXml = myPrevEle.InnerXml.Replace(" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                }

                //save changes to doc
                wDoc.Save();
            }
        }


        public static void checkForIndent(WordprocessingDocument doc, int e, string setIndentYN)
        {
            /* 
             *** Check for indents to add/remove when appropriate
             ***  - add at regular paragraphs, when missing
             ***  - remove at start of chapter, when found
            */
            var wDoc = doc.MainDocumentPart.Document;
            var indentCode = "<w:ind w:firstLine=\"720\" />";
            if (wDoc.Descendants<Paragraph>().ElementAtOrDefault(e) != null && wDoc.Descendants<Paragraph>().ElementAtOrDefault(e - 1) != null)
            {

                // Get CURRENT element details
                var myCurrEle = wDoc.Body.Descendants<Paragraph>().ElementAt(e);
                var myCurrEleOuterXml = myCurrEle.OuterXml;
                // T/F whether indent exists
                var indExistsTF = Regex.Match(myCurrEleOuterXml, indentCode).Success;
                // Start of chapter -- PREVIOUS element
                var myPrevEle = wDoc.Body.Descendants<Paragraph>().ElementAt(e - 1);
                var myPrevEleOuterXml = myPrevEle.OuterXml;
                var myPrevEleText = myPrevEle.InnerText;
                var startChapterTF = Regex.Match(myPrevEleText, "Chapter ").Success;                

                if (startChapterTF == true)
                {
                    // Start of chapter -- NEXT element -- No indent; remove if exists
                    if (indExistsTF == true)
                    {
                        // Remove indent else nada
                        myCurrEle.InnerXml = myCurrEle.InnerXml.Replace(indentCode, "");
                    }
                }
                else
                {
                    // Paragraphs in chapter -- CURRENT element -- Add indent if not there
                    if (indExistsTF == false)
                    {
                        // Insert indent else nada
                        myCurrEle.InnerXml = myCurrEle.InnerXml.Replace("</w:pPr>", indentCode + "</w:pPr>");
                    }
                }

                //cleanup
                myCurrEle.InnerXml = myCurrEle.InnerXml.Replace(" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                //save changes to doc
                wDoc.Save();
            }
        
        }


    }
}
