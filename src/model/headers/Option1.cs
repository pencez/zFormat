using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace zFormat.model.headers
{
    class Option1
    {
        // Adds child parts and generates content of the specified part.
        public static void CreateHeaderPart(HeaderPart part, int i)
        {
            if (i == 1)
            {
                GenerateTitleHeader(part);
            }
            else
            {
                GenerateAuthorHeader(part);
            }
        }

        // Generates TITLE content of part.
        private static void GenerateTitleHeader(HeaderPart part)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();

            RunProperties runProperties1 = new RunProperties();
            Color color1 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };

            runProperties1.Append(color1);
            SdtAlias sdtAlias1 = new SdtAlias() { Val = "Title" };
            Tag tag1 = new Tag() { Val = "" };
            SdtId sdtId1 = new SdtId() { Val = -162626310 };

            SdtPlaceholder sdtPlaceholder1 = new SdtPlaceholder();
            DocPartReference docPartReference1 = new DocPartReference() { Val = "C8E29D4D819B4064AF86E63E27701566" };

            sdtPlaceholder1.Append(docPartReference1);
            DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
            SdtContentText sdtContentText1 = new SdtContentText();

            sdtProperties1.Append(runProperties1);
            sdtProperties1.Append(sdtAlias1);
            sdtProperties1.Append(tag1);
            sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(sdtPlaceholder1);
            sdtProperties1.Append(dataBinding1);
            sdtProperties1.Append(sdtContentText1);

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "005A7282", RsidRunAdditionDefault = "005A7282" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Clear, Position = 4680 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Clear, Position = 9360 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            Justification justification1 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            Color color2 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };

            paragraphMarkRunProperties1.Append(color2);

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(tabs1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties2 = new RunProperties();
            Color color3 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };

            runProperties2.Append(color3);
            Text text1 = new Text();
            text1.Text = "GHOST OF SQUIRE HOUSE";

            run1.Append(runProperties2);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            sdtContentBlock1.Append(paragraph1);

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtContentBlock1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "005A7282", RsidRunAdditionDefault = "005A7282" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties2.Append(paragraphStyleId2);

            paragraph2.Append(paragraphProperties2);

            header1.Append(sdtBlock1);
            header1.Append(paragraph2);

            part.Header = header1;
        }


        // Generates AUTHOR content of part.
        private static void GenerateAuthorHeader(HeaderPart part)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "005A7282", RsidParagraphAddition = "005A7282", RsidParagraphProperties = "005A7282", RsidRunAdditionDefault = "005A7282" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Clear, Position = 4680 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Clear, Position = 9360 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            Justification justification1 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            Color color1 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };

            paragraphMarkRunProperties1.Append(color1);

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(tabs1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            Color color2 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };

            runProperties1.Append(color2);
            TabChar tabChar1 = new TabChar();

            run1.Append(runProperties1);
            run1.Append(tabChar1);

            SdtRun sdtRun1 = new SdtRun();

            SdtProperties sdtProperties1 = new SdtProperties();
            SdtAlias sdtAlias1 = new SdtAlias() { Val = "Author" };
            Tag tag1 = new Tag() { Val = "" };
            SdtId sdtId1 = new SdtId() { Val = 1923138210 };

            SdtPlaceholder sdtPlaceholder1 = new SdtPlaceholder();
            DocPartReference docPartReference1 = new DocPartReference() { Val = "50C0987FC8BD4240AED298B1EBAD60C2" };

            sdtPlaceholder1.Append(docPartReference1);
            DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:creator[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
            SdtContentText sdtContentText1 = new SdtContentText();

            sdtProperties1.Append(sdtAlias1);
            sdtProperties1.Append(tag1);
            sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(sdtPlaceholder1);
            sdtProperties1.Append(dataBinding1);
            sdtProperties1.Append(sdtContentText1);

            SdtContentRun sdtContentRun1 = new SdtContentRun();

            Run run2 = new Run();
            Text text1 = new Text();
            text1.Text = "Zach Pence";

            run2.Append(text1);

            sdtContentRun1.Append(run2);

            sdtRun1.Append(sdtProperties1);
            sdtRun1.Append(sdtContentRun1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(sdtRun1);

            header1.Append(paragraph1);

            part.Header = header1;
        }


    }
}
