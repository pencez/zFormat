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

                        var newFont = new RunFonts
                        {
                            Ascii = zFont,
                            EastAsia = zFont,
                            HighAnsi = zFont,
                            ComplexScript = zFont
                        };

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


        // Set the font for a chapter heading text run.
        public static void SetHeadFont(Run theRun, RunProperties subRunProp, String hFont)
        {
            var newFont = new RunFonts
            {
                Ascii = hFont,
                //newFont.EastAsia = hFont;
                HighAnsi = hFont
                //newFont.ComplexScript = hFont;
            };

            if (subRunProp != null)
            {
                var font = subRunProp.Descendants<RunFonts>().FirstOrDefault();
                subRunProp.ReplaceChild<RunFonts>(newFont, font);
            }
            else
            {
                var tmpSubRunProp = new RunProperties();
                tmpSubRunProp.AppendChild<RunFonts>(newFont);
                theRun.AppendChild<RunProperties>(tmpSubRunProp);
            }
        }


        // Set the font for chapter text headings.
        public static void SetTheChapterHeadingFont(FileInfo fileName, String hFont, String bFont, String dropCaseYN)
        {
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fileName.FullName, true))
            {
                Body body = wDoc.MainDocumentPart.Document.Body;
                //Get all paragraphs
                var lstParagraphs = body.Descendants<Paragraph>().ToList();
                int p = 0;
                var Chap1FoundYN = "No";
                var SetCh1RunStyle = "No";
                var pageBreakFoundYN = "No";
                foreach (var para in lstParagraphs)
                {
                    var theParaProps = para.Descendants<ParagraphProperties>().FirstOrDefault();
                    //Detect chapters by finding page breaks
                    if (theParaProps.PageBreakBefore != null)
                    {
                        pageBreakFoundYN = "Yes";
                    }
                    var subRuns = para.Descendants<Run>().ToList();
                    foreach (var run in subRuns)
                    {   
                        //runs once
                        if (SetCh1RunStyle == "Yes")
                        {
                            //Call special design options (Dropcap and ALLCAPS start run...)
                            SetDropCap(body, para, run, bFont);
                            SetCh1RunStyle = "No";
                        }

                        //Detect chapters by finding page breaks (secondary method)
                        //<w:br w:type="page"/>
                        var theBr = run.Descendants<Break>().FirstOrDefault();
                        if (theBr != null) {
                            if (theBr.Type == BreakValues.Page)
                            {
                                pageBreakFoundYN = "Yes";
                            }
                        }

                        
                        if (Chap1FoundYN == "No")
                        {
                            if (theParaProps.Justification != null)
                            {
                                //runs only one time
                                if (theParaProps.Justification.Val == JustificationValues.Center)
                                {
                                    //Might be chapter 1... Check for bold in rPr
                                    var theRunProps = run.Descendants<RunProperties>().FirstOrDefault();
                                    if (theRunProps.Bold != null)
                                    {
                                        //Bold was found too, is there text in the run?
                                        if (run.InnerText.Length > 2) {
                                            Chap1FoundYN = "Yes";
                                            SetHeadFont(run, theRunProps, hFont);
                                            SetCh1RunStyle = "Yes";
                                        }
                                    }
                                }
                            }
                        } else
                        {
                            if (pageBreakFoundYN == "Yes") {
                                if (theParaProps.Justification != null)
                                {
                                    if (theParaProps.Justification.Val == JustificationValues.Center)
                                    {
                                        //Check for bold in rPr
                                        var theRunProps = run.Descendants<RunProperties>().FirstOrDefault();
                                        if (theRunProps.Bold != null)
                                        {
                                            //Bold was found too, is there text in the run?
                                            if (run.InnerText.Length > 2)
                                            {
                                                SetHeadFont(run, theRunProps, hFont);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    //New chapter text starting; heading style ends
                                    pageBreakFoundYN = "No";
                                    //Call special design options (Dropcap and ALLCAPS start run...)
                                    SetDropCap(body, para, run, bFont);
                                }

                            }
                            
                        }

                    }
                    p++;
                }
                wDoc.MainDocumentPart.Document.Save();
                wDoc.Close();
            }
        }


        // Set the design for subchapter marking
        public static void SetSubChapterStyling(FileInfo fileName, String design)
        {
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fileName.FullName, true))
            {
                Body body = wDoc.MainDocumentPart.Document.Body;
                //Get all paragraphs
                var lstParagraphs = body.Descendants<Paragraph>().ToList();
                int p = 0;
                foreach (var para in lstParagraphs)
                {
                    var theParaProps = para.Descendants<ParagraphProperties>().FirstOrDefault();
                    var subParaText = para.Descendants<Text>().FirstOrDefault();
                    if (theParaProps != null && subParaText != null)
                    {
                        if (theParaProps.Justification != null)
                        {
                            if (theParaProps.Justification.Val == JustificationValues.Center)
                            {
                                if (subParaText.Text == "*" || subParaText.Text == "**" || subParaText.Text == "***" 
                                    || subParaText.Text == "* *" || subParaText.Text == "* * *")
                                {
                                    subParaText.Text = "*" + design + "*";
                                }
                                if (subParaText.Text == "#" || subParaText.Text == "##" || subParaText.Text == "###" 
                                    || subParaText.Text == "# #" || subParaText.Text == "# # #")
                                {
                                    subParaText.Text = "#" + design + "#";
                                }
                            }
                        }
                    }
                    p++;
                }
            }
        }


        // Set Drop Cap for first character in chapter paragraph
        public static void SetDropCap(Body body, Paragraph para, Run run, String bFont)
        {
            // Removes the indent from first line in paragraph
            var subParaProp = para.Descendants<ParagraphProperties>().FirstOrDefault();
            if (subParaProp.Indentation != null)
            {
                subParaProp.Indentation.Remove();

                // Store the first character of paragraph, for drop cap
                var getfirstLetter = run.InnerText.Substring(0, 1);

                // Remove first letter from paragraph
                var zNewString = run.LastChild.InnerText.Substring(1);
                var subParaText = para.Descendants<Text>().FirstOrDefault();
                subParaText.Text = zNewString;

                //Create new para + run for the drop cap letter
                Paragraph newP = new Paragraph();
                ParagraphProperties new_pPr = new ParagraphProperties();
                FrameProperties newFrP = new FrameProperties
                {
                    DropCap = DropCapLocationValues.Drop,
                    Lines = 3,
                    Wrap = TextWrappingValues.Around,
                    VerticalPosition = VerticalAnchorValues.Text,
                    HorizontalPosition = HorizontalAnchorValues.Text
                };
                new_pPr.AppendChild<FrameProperties>(newFrP);
                SpacingBetweenLines newS = new SpacingBetweenLines
                {
                    Line = "1076",
                    LineRule = LineSpacingRuleValues.Exact
                };
                new_pPr.AppendChild<SpacingBetweenLines>(newS);
                TextAlignment newTA = new TextAlignment
                {
                    Val = VerticalTextAlignmentValues.Baseline
                };
                new_pPr.AppendChild<TextAlignment>(newTA);

                RunProperties new_rPr = new RunProperties();
                RunFonts rFonts = new RunFonts()
                {
                    Ascii = bFont,
                    EastAsia = bFont,
                    HighAnsi = bFont,
                    ComplexScript = bFont
                };
                new_rPr.AppendChild<RunFonts>(rFonts);
                Position newPos = new Position()
                {
                    Val = "-8"
                };
                new_rPr.AppendChild<Position>(newPos);
                FontSize newSize = new FontSize()
                {
                    Val = "134"
                };
                new_rPr.AppendChild<FontSize>(newSize);
                FontSizeComplexScript newCsSize = new FontSizeComplexScript()
                {
                    Val = "24"
                };
                new_rPr.AppendChild<FontSizeComplexScript>(newCsSize);

                //now connect everything together; the <w:p> element
                Run newR = new Run();
                Text newT = new Text(getfirstLetter);
                newR.Append(new_rPr);
                newR.Append(newT);
                newP.Append(new_pPr);
                newP.Append(newR);

                //insert to previous para
                body.InsertBefore(newP, para);
            }
        }


    }
}
