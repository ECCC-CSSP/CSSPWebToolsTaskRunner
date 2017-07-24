using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using System.Linq;
using System;

namespace CSSPWebToolsTaskRunner.Services
{
   public partial  class GeneratedClassFCFormDocx
    {
        private void CreateWaterAnalysisDataSheet(Body body, int SheetNumber)
        {
            List<string> AllowableTideTextList = new List<string>()
            {
                "LT", "LR", "LF", "MT", "MR", "MF", "HT", "HR", "HF"
            };

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00B67DA8", RsidParagraphAddition = "009A75C9", RsidParagraphProperties = "009A75C9", RsidRunAdditionDefault = "006D4BC2" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold1 = new Bold();

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(bold1);

            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run() { RsidRunProperties = "0005391D" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold2 = new Bold();

            runProperties1.Append(runFonts2);
            runProperties1.Append(bold2);
            Text text1 = new Text();
            text1.Text = "A1 Fecal Coliform- Water Analysis Data Sheet";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "15120", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = -72, Type = TableWidthUnitValues.Dxa };
            TableLook tableLook1 = new TableLook() { Val = "01E0" };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1364" };
            GridColumn gridColumn2 = new GridColumn() { Width = "4015" };
            GridColumn gridColumn3 = new GridColumn() { Width = "236" };
            GridColumn gridColumn4 = new GridColumn() { Width = "1191" };
            GridColumn gridColumn5 = new GridColumn() { Width = "4012" };
            GridColumn gridColumn6 = new GridColumn() { Width = "236" };
            GridColumn gridColumn7 = new GridColumn() { Width = "845" };
            GridColumn gridColumn8 = new GridColumn() { Width = "3221" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);
            tableGrid1.Append(gridColumn6);
            tableGrid1.Append(gridColumn7);
            tableGrid1.Append(gridColumn8);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009677FE", RsidTableRowProperties = "00ED7FB2" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "1364", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder1);
            tableCellBorders1.Append(leftBorder1);
            tableCellBorders1.Append(bottomBorder1);
            tableCellBorders1.Append(rightBorder1);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(shading1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009677FE", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "009677FE" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties2.Append(runFonts9);

            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run8 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties8.Append(runFonts10);
            Text text8 = new Text();
            text8.Text = "Subsector:";

            run8.Append(runProperties8);
            run8.Append(text8);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run8);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "4015", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder2);
            tableCellBorders2.Append(leftBorder2);
            tableCellBorders2.Append(bottomBorder2);
            tableCellBorders2.Append(rightBorder2);
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(tableCellBorders2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009677FE", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties3.Append(runFonts11);

            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties9.Append(runFonts12);
            Text text9 = new Text();
            text9.Text = labSheetA1Sheet.SubsectorName.Substring(0, labSheetA1Sheet.SubsectorName.IndexOf(" "));

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run9);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph3);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(leftBorder3);
            tableCellBorders3.Append(bottomBorder3);
            tableCellBorders3.Append(rightBorder3);
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellBorders3);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellVerticalAlignment3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009677FE", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "009677FE" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties4.Append(runFonts13);

            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            paragraph4.Append(paragraphProperties4);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph4);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1191", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(topBorder3);
            tableCellBorders4.Append(leftBorder4);
            tableCellBorders4.Append(bottomBorder4);
            tableCellBorders4.Append(rightBorder4);
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(tableCellBorders4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellVerticalAlignment4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009677FE", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "009677FE" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties5.Append(runFonts14);

            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run10 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties10.Append(runFonts15);
            Text text10 = new Text();
            text10.Text = "Location:";

            run10.Append(runProperties10);
            run10.Append(text10);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run10);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "4012", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(topBorder4);
            tableCellBorders5.Append(leftBorder5);
            tableCellBorders5.Append(bottomBorder5);
            tableCellBorders5.Append(rightBorder5);
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(tableCellBorders5);
            tableCellProperties5.Append(shading5);
            tableCellProperties5.Append(tableCellVerticalAlignment5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009677FE", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties6.Append(runFonts16);

            paragraphProperties6.Append(justification6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties11.Append(runFonts17);
            Text text11 = new Text();
            string Location = labSheetA1Sheet.SubsectorName.Substring(labSheetA1Sheet.SubsectorName.IndexOf(" ") + 1);
            text11.Text = (Location.Length > 35 ? Location.Substring(0, 30) + " ..." : Location);

            run11.Append(runProperties11);
            run11.Append(text11);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run11);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph6);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(leftBorder6);
            tableCellBorders6.Append(bottomBorder6);
            tableCellBorders6.Append(rightBorder6);
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(tableCellBorders6);
            tableCellProperties6.Append(shading6);
            tableCellProperties6.Append(tableCellVerticalAlignment6);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009677FE", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "009677FE" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            Justification justification7 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties7.Append(runFonts18);

            paragraphProperties7.Append(justification7);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            paragraph7.Append(paragraphProperties7);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph7);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "845", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders7 = new TableCellBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders7.Append(topBorder5);
            tableCellBorders7.Append(leftBorder7);
            tableCellBorders7.Append(bottomBorder7);
            tableCellBorders7.Append(rightBorder7);
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(tableCellBorders7);
            tableCellProperties7.Append(shading7);
            tableCellProperties7.Append(tableCellVerticalAlignment7);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009677FE", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "009677FE" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            Justification justification8 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties8.Append(runFonts19);

            paragraphProperties8.Append(justification8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            Run run12 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties12.Append(runFonts20);
            Text text12 = new Text();
            text12.Text = "Date:";

            run12.Append(runProperties12);
            run12.Append(text12);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run12);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph8);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "3221", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders8 = new TableCellBorders();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders8.Append(topBorder6);
            tableCellBorders8.Append(leftBorder8);
            tableCellBorders8.Append(bottomBorder8);
            tableCellBorders8.Append(rightBorder8);
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellBorders8);
            tableCellProperties8.Append(shading8);
            tableCellProperties8.Append(tableCellVerticalAlignment8);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009677FE", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            Justification justification9 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color1 = new Color() { Val = "EAEAEA" };

            paragraphMarkRunProperties9.Append(runFonts21);
            paragraphMarkRunProperties9.Append(color1);

            paragraphProperties9.Append(justification9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run13 = new Run();

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties13.Append(runFonts22);
            Text text13 = new Text();
            DateTime SampleDate = new DateTime(int.Parse(labSheetA1Sheet.RunYear), int.Parse(labSheetA1Sheet.RunMonth), int.Parse(labSheetA1Sheet.RunDay));
            text13.Text = SampleDate.ToString("yyyy MMMM dd");


            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run13);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph9);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);
            tableRow1.Append(tableCell5);
            tableRow1.Append(tableCell6);
            tableRow1.Append(tableCell7);
            tableRow1.Append(tableCell8);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "00004340", RsidRunAdditionDefault = "00004340" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties10.Append(runFonts25);

            paragraphProperties10.Append(paragraphMarkRunProperties10);

            paragraph10.Append(paragraphProperties10);

            Table table2 = new Table();

            TableProperties tableProperties2 = new TableProperties();
            TableWidth tableWidth2 = new TableWidth() { Width = "15120", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation2 = new TableIndentation() { Width = -72, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder7);
            tableBorders1.Append(leftBorder9);
            tableBorders1.Append(bottomBorder9);
            tableBorders1.Append(rightBorder9);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLayout tableLayout1 = new TableLayout() { Type = TableLayoutValues.Fixed };
            TableLook tableLook2 = new TableLook() { Val = "01E0" };

            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableIndentation2);
            tableProperties2.Append(tableBorders1);
            tableProperties2.Append(tableLayout1);
            tableProperties2.Append(tableLook2);

            TableGrid tableGrid2 = new TableGrid();
            GridColumn gridColumn9 = new GridColumn() { Width = "1260" };
            GridColumn gridColumn10 = new GridColumn() { Width = "1800" };
            GridColumn gridColumn11 = new GridColumn() { Width = "236" };
            GridColumn gridColumn12 = new GridColumn() { Width = "2284" };
            GridColumn gridColumn13 = new GridColumn() { Width = "720" };
            GridColumn gridColumn14 = new GridColumn() { Width = "236" };
            GridColumn gridColumn15 = new GridColumn() { Width = "1024" };
            GridColumn gridColumn16 = new GridColumn() { Width = "236" };
            GridColumn gridColumn17 = new GridColumn() { Width = "1024" };
            GridColumn gridColumn18 = new GridColumn() { Width = "900" };
            GridColumn gridColumn19 = new GridColumn() { Width = "1080" };
            GridColumn gridColumn20 = new GridColumn() { Width = "900" };
            GridColumn gridColumn21 = new GridColumn() { Width = "900" };
            GridColumn gridColumn22 = new GridColumn() { Width = "900" };
            GridColumn gridColumn23 = new GridColumn() { Width = "1620" };

            tableGrid2.Append(gridColumn9);
            tableGrid2.Append(gridColumn10);
            tableGrid2.Append(gridColumn11);
            tableGrid2.Append(gridColumn12);
            tableGrid2.Append(gridColumn13);
            tableGrid2.Append(gridColumn14);
            tableGrid2.Append(gridColumn15);
            tableGrid2.Append(gridColumn16);
            tableGrid2.Append(gridColumn17);
            tableGrid2.Append(gridColumn18);
            tableGrid2.Append(gridColumn19);
            tableGrid2.Append(gridColumn20);
            tableGrid2.Append(gridColumn21);
            tableGrid2.Append(gridColumn22);
            tableGrid2.Append(gridColumn23);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00397699", RsidTableRowProperties = "00ED7FB2" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)230U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders9 = new TableCellBorders();
            TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders9.Append(topBorder8);
            tableCellBorders9.Append(leftBorder10);
            tableCellBorders9.Append(rightBorder10);
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(verticalMerge1);
            tableCellProperties9.Append(tableCellBorders9);
            tableCellProperties9.Append(shading9);
            tableCellProperties9.Append(tableCellVerticalAlignment9);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties11.Append(runFonts26);
            paragraphMarkRunProperties11.Append(fontSize1);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript1);

            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run16 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize2 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "20" };

            runProperties16.Append(runFonts27);
            runProperties16.Append(fontSize2);
            runProperties16.Append(fontSizeComplexScript2);
            Text text16 = new Text();
            text16.Text = "Sample Time";

            run16.Append(runProperties16);
            run16.Append(text16);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run16);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph11);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "1800", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge2 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders10 = new TableCellBorders();
            TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder11 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder11 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders10.Append(topBorder9);
            tableCellBorders10.Append(leftBorder11);
            tableCellBorders10.Append(rightBorder11);
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(verticalMerge2);
            tableCellProperties10.Append(tableCellBorders10);
            tableCellProperties10.Append(shading10);
            tableCellProperties10.Append(tableCellVerticalAlignment10);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color2 = new Color() { Val = "C0C0C0" };
            FontSize fontSize3 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties12.Append(runFonts28);
            paragraphMarkRunProperties12.Append(color2);
            paragraphMarkRunProperties12.Append(fontSize3);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript3);

            paragraphProperties12.Append(justification10);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run17 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize4 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "20" };

            runProperties17.Append(runFonts29);
            runProperties17.Append(fontSize4);
            runProperties17.Append(fontSizeComplexScript4);
            Text text17 = new Text();
            text17.Text = labSheetA1Sheet.LabSheetA1MeasurementList.Min(c => c.Time) + "/" + labSheetA1Sheet.LabSheetA1MeasurementList.Max(c => c.Time);

            run17.Append(runProperties17);
            run17.Append(text17);

            Run run18 = new Run() { RsidRunProperties = "0072154A" };

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run17);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph12);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge3 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders11 = new TableCellBorders();
            TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder12 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder12 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders11.Append(topBorder10);
            tableCellBorders11.Append(leftBorder12);
            tableCellBorders11.Append(rightBorder12);
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment11 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(verticalMerge3);
            tableCellProperties11.Append(tableCellBorders11);
            tableCellProperties11.Append(shading11);
            tableCellProperties11.Append(tableCellVerticalAlignment11);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize7 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties13.Append(runFonts32);
            paragraphMarkRunProperties13.Append(fontSize7);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript7);

            paragraphProperties13.Append(paragraphMarkRunProperties13);

            paragraph13.Append(paragraphProperties13);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph13);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "2284", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge4 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders12 = new TableCellBorders();
            TopBorder topBorder11 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder13 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder13 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders12.Append(topBorder11);
            tableCellBorders12.Append(leftBorder13);
            tableCellBorders12.Append(rightBorder13);
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(verticalMerge4);
            tableCellProperties12.Append(tableCellBorders12);
            tableCellProperties12.Append(shading12);
            tableCellProperties12.Append(tableCellVerticalAlignment12);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize8 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties14.Append(runFonts33);
            paragraphMarkRunProperties14.Append(fontSize8);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript8);

            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run20 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize9 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "20" };

            runProperties20.Append(runFonts34);
            runProperties20.Append(fontSize9);
            runProperties20.Append(fontSizeComplexScript9);
            Text text20 = new Text();
            text20.Text = "Incubation Start Time";

            run20.Append(runProperties20);
            run20.Append(text20);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run20);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph14);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge5 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders13 = new TableCellBorders();
            TopBorder topBorder12 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder14 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder14 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders13.Append(topBorder12);
            tableCellBorders13.Append(leftBorder14);
            tableCellBorders13.Append(rightBorder14);
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment13 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(verticalMerge5);
            tableCellProperties13.Append(tableCellBorders13);
            tableCellProperties13.Append(shading13);
            tableCellProperties13.Append(tableCellVerticalAlignment13);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts35 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize10 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties15.Append(runFonts35);
            paragraphMarkRunProperties15.Append(fontSize10);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript10);

            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run21 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize11 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "20" };

            runProperties21.Append(runFonts36);
            runProperties21.Append(fontSize11);
            runProperties21.Append(fontSizeComplexScript11);
            Text text21 = new Text();
            text21.Text = labSheetA1Sheet.IncubationStartTime;

            run21.Append(runProperties21);
            run21.Append(text21);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run21);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph15);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge6 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders14 = new TableCellBorders();
            TopBorder topBorder13 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder15 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder15 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders14.Append(topBorder13);
            tableCellBorders14.Append(leftBorder15);
            tableCellBorders14.Append(rightBorder15);
            Shading shading14 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment14 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties14.Append(tableCellWidth14);
            tableCellProperties14.Append(verticalMerge6);
            tableCellProperties14.Append(tableCellBorders14);
            tableCellProperties14.Append(shading14);
            tableCellProperties14.Append(tableCellVerticalAlignment14);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize14 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties16.Append(runFonts39);
            paragraphMarkRunProperties16.Append(fontSize14);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript14);

            paragraphProperties16.Append(paragraphMarkRunProperties16);

            paragraph16.Append(paragraphProperties16);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph16);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge7 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders15 = new TableCellBorders();
            TopBorder topBorder14 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder16 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder16 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders15.Append(topBorder14);
            tableCellBorders15.Append(leftBorder16);
            tableCellBorders15.Append(rightBorder16);
            Shading shading15 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment15 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(verticalMerge7);
            tableCellProperties15.Append(tableCellBorders15);
            tableCellProperties15.Append(shading15);
            tableCellProperties15.Append(tableCellVerticalAlignment15);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize15 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties17.Append(runFonts40);
            paragraphMarkRunProperties17.Append(fontSize15);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript15);

            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run24 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize16 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "20" };

            runProperties24.Append(runFonts41);
            runProperties24.Append(fontSize16);
            runProperties24.Append(fontSizeComplexScript16);
            Text text24 = new Text();
            text24.Text = "TC (";

            run24.Append(runProperties24);
            run24.Append(text24);

            Run run25 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize17 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "20" };
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties25.Append(runFonts42);
            runProperties25.Append(fontSize17);
            runProperties25.Append(fontSizeComplexScript17);
            runProperties25.Append(verticalTextAlignment1);
            Text text25 = new Text();
            text25.Text = "o";

            run25.Append(runProperties25);
            run25.Append(text25);

            Run run26 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize18 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "20" };

            runProperties26.Append(runFonts43);
            runProperties26.Append(fontSize18);
            runProperties26.Append(fontSizeComplexScript18);
            Text text26 = new Text();
            text26.Text = "C)";

            run26.Append(runProperties26);
            run26.Append(text26);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run24);
            paragraph17.Append(run25);
            paragraph17.Append(run26);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph17);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge8 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders16 = new TableCellBorders();
            TopBorder topBorder15 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder17 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder17 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders16.Append(topBorder15);
            tableCellBorders16.Append(leftBorder17);
            tableCellBorders16.Append(rightBorder17);
            Shading shading16 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment16 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(verticalMerge8);
            tableCellProperties16.Append(tableCellBorders16);
            tableCellProperties16.Append(shading16);
            tableCellProperties16.Append(tableCellVerticalAlignment16);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize19 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties18.Append(runFonts44);
            paragraphMarkRunProperties18.Append(fontSize19);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript19);

            paragraphProperties18.Append(paragraphMarkRunProperties18);

            paragraph18.Append(paragraphProperties18);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph18);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders17 = new TableCellBorders();
            TopBorder topBorder16 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder18 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder18 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders17.Append(topBorder16);
            tableCellBorders17.Append(leftBorder18);
            tableCellBorders17.Append(bottomBorder10);
            tableCellBorders17.Append(rightBorder18);
            Shading shading17 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment17 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties17.Append(tableCellWidth17);
            tableCellProperties17.Append(tableCellBorders17);
            tableCellProperties17.Append(shading17);
            tableCellProperties17.Append(tableCellVerticalAlignment17);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold9 = new Bold();
            FontSize fontSize20 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties19.Append(runFonts45);
            paragraphMarkRunProperties19.Append(bold9);
            paragraphMarkRunProperties19.Append(fontSize20);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript20);

            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run27 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold10 = new Bold();
            FontSize fontSize21 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "20" };

            runProperties27.Append(runFonts46);
            runProperties27.Append(bold10);
            runProperties27.Append(fontSize21);
            runProperties27.Append(fontSizeComplexScript21);
            Text text27 = new Text();
            text27.Text = "Control";

            run27.Append(runProperties27);
            run27.Append(text27);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run27);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph19);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders18 = new TableCellBorders();
            TopBorder topBorder17 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder19 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder19 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders18.Append(topBorder17);
            tableCellBorders18.Append(leftBorder19);
            tableCellBorders18.Append(rightBorder19);
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment18 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellBorders18);
            tableCellProperties18.Append(shading18);
            tableCellProperties18.Append(tableCellVerticalAlignment18);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic1 = new Italic();
            FontSize fontSize22 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties20.Append(runFonts47);
            paragraphMarkRunProperties20.Append(italic1);
            paragraphMarkRunProperties20.Append(fontSize22);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript22);

            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run28 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts48 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic2 = new Italic();
            FontSize fontSize23 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "18" };

            runProperties28.Append(runFonts48);
            runProperties28.Append(italic2);
            runProperties28.Append(fontSize23);
            runProperties28.Append(fontSizeComplexScript23);
            Text text28 = new Text();
            text28.Text = "E. coli";

            run28.Append(runProperties28);
            run28.Append(text28);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run28);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph20);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders19 = new TableCellBorders();
            TopBorder topBorder18 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder20 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder20 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders19.Append(topBorder18);
            tableCellBorders19.Append(leftBorder20);
            tableCellBorders19.Append(rightBorder20);
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment19 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(tableCellBorders19);
            tableCellProperties19.Append(shading19);
            tableCellProperties19.Append(tableCellVerticalAlignment19);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize24 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties21.Append(runFonts49);
            paragraphMarkRunProperties21.Append(fontSize24);
            paragraphMarkRunProperties21.Append(fontSizeComplexScript24);

            paragraphProperties21.Append(paragraphMarkRunProperties21);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run29 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic3 = new Italic();
            FontSize fontSize25 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "18" };

            runProperties29.Append(runFonts50);
            runProperties29.Append(italic3);
            runProperties29.Append(fontSize25);
            runProperties29.Append(fontSizeComplexScript25);
            Text text29 = new Text();
            text29.Text = "E.cloacae";

            run29.Append(runProperties29);
            run29.Append(text29);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(proofError1);
            paragraph21.Append(run29);
            paragraph21.Append(proofError2);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph21);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders20 = new TableCellBorders();
            TopBorder topBorder19 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder21 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder21 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders20.Append(topBorder19);
            tableCellBorders20.Append(leftBorder21);
            tableCellBorders20.Append(rightBorder21);
            Shading shading20 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment20 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(tableCellBorders20);
            tableCellProperties20.Append(shading20);
            tableCellProperties20.Append(tableCellVerticalAlignment20);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts51 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic4 = new Italic();
            FontSize fontSize26 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties22.Append(runFonts51);
            paragraphMarkRunProperties22.Append(italic4);
            paragraphMarkRunProperties22.Append(fontSize26);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript26);

            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run30 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic5 = new Italic();
            FontSize fontSize27 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "18" };

            runProperties30.Append(runFonts52);
            runProperties30.Append(italic5);
            runProperties30.Append(fontSize27);
            runProperties30.Append(fontSizeComplexScript27);
            Text text30 = new Text();
            text30.Text = "S. epi";

            run30.Append(runProperties30);
            run30.Append(text30);

            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(run30);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph22);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge9 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders21 = new TableCellBorders();
            TopBorder topBorder20 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder22 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder22 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders21.Append(topBorder20);
            tableCellBorders21.Append(leftBorder22);
            tableCellBorders21.Append(rightBorder22);
            Shading shading21 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment21 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(verticalMerge9);
            tableCellProperties21.Append(tableCellBorders21);
            tableCellProperties21.Append(shading21);
            tableCellProperties21.Append(tableCellVerticalAlignment21);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts53 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize28 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties23.Append(runFonts53);
            paragraphMarkRunProperties23.Append(fontSize28);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript28);

            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run31 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts54 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize29 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "20" };

            runProperties31.Append(runFonts54);
            runProperties31.Append(fontSize29);
            runProperties31.Append(fontSizeComplexScript29);
            Text text31 = new Text();
            text31.Text = "Blank";

            run31.Append(runProperties31);
            run31.Append(text31);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(run31);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph23);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge10 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders22 = new TableCellBorders();
            TopBorder topBorder21 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder23 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder23 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders22.Append(topBorder21);
            tableCellBorders22.Append(leftBorder23);
            tableCellBorders22.Append(rightBorder23);
            Shading shading22 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment22 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(verticalMerge10);
            tableCellProperties22.Append(tableCellBorders22);
            tableCellProperties22.Append(shading22);
            tableCellProperties22.Append(tableCellVerticalAlignment22);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts55 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize30 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties24.Append(runFonts55);
            paragraphMarkRunProperties24.Append(fontSize30);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript30);

            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run32 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts56 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize31 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "20" };

            runProperties32.Append(runFonts56);
            runProperties32.Append(fontSize31);
            runProperties32.Append(fontSizeComplexScript31);
            Text text32 = new Text();
            text32.Text = "A1";

            run32.Append(runProperties32);
            run32.Append(text32);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run32);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize32 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties25.Append(runFonts57);
            paragraphMarkRunProperties25.Append(fontSize32);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript32);

            paragraphProperties25.Append(paragraphMarkRunProperties25);

            Run run33 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts58 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize33 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "20" };

            runProperties33.Append(runFonts58);
            runProperties33.Append(fontSize33);
            runProperties33.Append(fontSizeComplexScript33);
            Text text33 = new Text();
            text33.Text = "Media";

            run33.Append(runProperties33);
            run33.Append(text33);

            paragraph25.Append(paragraphProperties25);
            paragraph25.Append(run33);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph24);
            tableCell22.Append(paragraph25);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge11 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders23 = new TableCellBorders();
            TopBorder topBorder22 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder24 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder24 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders23.Append(topBorder22);
            tableCellBorders23.Append(leftBorder24);
            tableCellBorders23.Append(rightBorder24);
            Shading shading23 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment23 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(verticalMerge11);
            tableCellProperties23.Append(tableCellBorders23);
            tableCellProperties23.Append(shading23);
            tableCellProperties23.Append(tableCellVerticalAlignment23);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts59 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize34 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties26.Append(runFonts59);
            paragraphMarkRunProperties26.Append(fontSize34);
            paragraphMarkRunProperties26.Append(fontSizeComplexScript34);

            paragraphProperties26.Append(paragraphMarkRunProperties26);

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<w:smartTag w:uri=\"urn:schemas-microsoft-com:office:smarttags\" w:element=\"place\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r w:rsidRPr=\"00ED7FB2\"><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\" /><w:sz w:val=\"20\" /><w:szCs w:val=\"20\" /></w:rPr><w:t>Lot</w:t></w:r></w:smartTag>");

            Run run34 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts60 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize35 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "20" };

            runProperties34.Append(runFonts60);
            runProperties34.Append(fontSize35);
            runProperties34.Append(fontSizeComplexScript35);
            Text text34 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text34.Text = " #";

            run34.Append(runProperties34);
            run34.Append(text34);

            Run run35 = new Run() { RsidRunProperties = "00ED7FB2", RsidRunAddition = "009A609A" };

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts61 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize36 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "20" };

            runProperties35.Append(runFonts61);
            runProperties35.Append(fontSize36);
            runProperties35.Append(fontSizeComplexScript36);
            Text text35 = new Text();
            text35.Text = "’s";

            run35.Append(runProperties35);
            run35.Append(text35);

            paragraph26.Append(paragraphProperties26);
            paragraph26.Append(openXmlUnknownElement1);
            paragraph26.Append(run34);
            paragraph26.Append(run35);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph26);

            tableRow2.Append(tableRowProperties1);
            tableRow2.Append(tableCell9);
            tableRow2.Append(tableCell10);
            tableRow2.Append(tableCell11);
            tableRow2.Append(tableCell12);
            tableRow2.Append(tableCell13);
            tableRow2.Append(tableCell14);
            tableRow2.Append(tableCell15);
            tableRow2.Append(tableCell16);
            tableRow2.Append(tableCell17);
            tableRow2.Append(tableCell18);
            tableRow2.Append(tableCell19);
            tableRow2.Append(tableCell20);
            tableRow2.Append(tableCell21);
            tableRow2.Append(tableCell22);
            tableRow2.Append(tableCell23);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "003511D5", RsidTableRowProperties = "00ED7FB2" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)230U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge12 = new VerticalMerge();

            TableCellBorders tableCellBorders24 = new TableCellBorders();
            LeftBorder leftBorder25 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder25 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders24.Append(leftBorder25);
            tableCellBorders24.Append(bottomBorder11);
            tableCellBorders24.Append(rightBorder25);
            Shading shading24 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment24 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties24.Append(tableCellWidth24);
            tableCellProperties24.Append(verticalMerge12);
            tableCellProperties24.Append(tableCellBorders24);
            tableCellProperties24.Append(shading24);
            tableCellProperties24.Append(tableCellVerticalAlignment24);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts62 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize37 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties27.Append(runFonts62);
            paragraphMarkRunProperties27.Append(fontSize37);
            paragraphMarkRunProperties27.Append(fontSizeComplexScript37);

            paragraphProperties27.Append(paragraphMarkRunProperties27);

            paragraph27.Append(paragraphProperties27);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph27);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "1800", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge13 = new VerticalMerge();

            TableCellBorders tableCellBorders25 = new TableCellBorders();
            LeftBorder leftBorder26 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder12 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder26 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders25.Append(leftBorder26);
            tableCellBorders25.Append(bottomBorder12);
            tableCellBorders25.Append(rightBorder26);
            Shading shading25 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment25 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties25.Append(tableCellWidth25);
            tableCellProperties25.Append(verticalMerge13);
            tableCellProperties25.Append(tableCellBorders25);
            tableCellProperties25.Append(shading25);
            tableCellProperties25.Append(tableCellVerticalAlignment25);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts63 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color3 = new Color() { Val = "C0C0C0" };
            FontSize fontSize38 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties28.Append(runFonts63);
            paragraphMarkRunProperties28.Append(color3);
            paragraphMarkRunProperties28.Append(fontSize38);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript38);

            paragraphProperties28.Append(paragraphMarkRunProperties28);

            paragraph28.Append(paragraphProperties28);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph28);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge14 = new VerticalMerge();

            TableCellBorders tableCellBorders26 = new TableCellBorders();
            LeftBorder leftBorder27 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder13 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder27 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders26.Append(leftBorder27);
            tableCellBorders26.Append(bottomBorder13);
            tableCellBorders26.Append(rightBorder27);
            Shading shading26 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment26 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties26.Append(tableCellWidth26);
            tableCellProperties26.Append(verticalMerge14);
            tableCellProperties26.Append(tableCellBorders26);
            tableCellProperties26.Append(shading26);
            tableCellProperties26.Append(tableCellVerticalAlignment26);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            RunFonts runFonts64 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize39 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties29.Append(runFonts64);
            paragraphMarkRunProperties29.Append(fontSize39);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript39);

            paragraphProperties29.Append(paragraphMarkRunProperties29);

            paragraph29.Append(paragraphProperties29);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph29);

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "2284", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge15 = new VerticalMerge();

            TableCellBorders tableCellBorders27 = new TableCellBorders();
            LeftBorder leftBorder28 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder14 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder28 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders27.Append(leftBorder28);
            tableCellBorders27.Append(bottomBorder14);
            tableCellBorders27.Append(rightBorder28);
            Shading shading27 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment27 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties27.Append(tableCellWidth27);
            tableCellProperties27.Append(verticalMerge15);
            tableCellProperties27.Append(tableCellBorders27);
            tableCellProperties27.Append(shading27);
            tableCellProperties27.Append(tableCellVerticalAlignment27);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            RunFonts runFonts65 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize40 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties30.Append(runFonts65);
            paragraphMarkRunProperties30.Append(fontSize40);
            paragraphMarkRunProperties30.Append(fontSizeComplexScript40);

            paragraphProperties30.Append(paragraphMarkRunProperties30);

            paragraph30.Append(paragraphProperties30);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph30);

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge16 = new VerticalMerge();

            TableCellBorders tableCellBorders28 = new TableCellBorders();
            LeftBorder leftBorder29 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder15 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder29 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders28.Append(leftBorder29);
            tableCellBorders28.Append(bottomBorder15);
            tableCellBorders28.Append(rightBorder29);
            Shading shading28 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment28 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties28.Append(tableCellWidth28);
            tableCellProperties28.Append(verticalMerge16);
            tableCellProperties28.Append(tableCellBorders28);
            tableCellProperties28.Append(shading28);
            tableCellProperties28.Append(tableCellVerticalAlignment28);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts66 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize41 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties31.Append(runFonts66);
            paragraphMarkRunProperties31.Append(fontSize41);
            paragraphMarkRunProperties31.Append(fontSizeComplexScript41);

            paragraphProperties31.Append(paragraphMarkRunProperties31);

            paragraph31.Append(paragraphProperties31);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph31);

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge17 = new VerticalMerge();

            TableCellBorders tableCellBorders29 = new TableCellBorders();
            LeftBorder leftBorder30 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder16 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder30 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders29.Append(leftBorder30);
            tableCellBorders29.Append(bottomBorder16);
            tableCellBorders29.Append(rightBorder30);
            Shading shading29 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment29 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties29.Append(tableCellWidth29);
            tableCellProperties29.Append(verticalMerge17);
            tableCellProperties29.Append(tableCellBorders29);
            tableCellProperties29.Append(shading29);
            tableCellProperties29.Append(tableCellVerticalAlignment29);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize42 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties32.Append(runFonts67);
            paragraphMarkRunProperties32.Append(fontSize42);
            paragraphMarkRunProperties32.Append(fontSizeComplexScript42);

            paragraphProperties32.Append(paragraphMarkRunProperties32);

            paragraph32.Append(paragraphProperties32);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph32);

            TableCell tableCell30 = new TableCell();

            TableCellProperties tableCellProperties30 = new TableCellProperties();
            TableCellWidth tableCellWidth30 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge18 = new VerticalMerge();

            TableCellBorders tableCellBorders30 = new TableCellBorders();
            LeftBorder leftBorder31 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder17 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder31 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders30.Append(leftBorder31);
            tableCellBorders30.Append(bottomBorder17);
            tableCellBorders30.Append(rightBorder31);
            Shading shading30 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment30 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties30.Append(tableCellWidth30);
            tableCellProperties30.Append(verticalMerge18);
            tableCellProperties30.Append(tableCellBorders30);
            tableCellProperties30.Append(shading30);
            tableCellProperties30.Append(tableCellVerticalAlignment30);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize43 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties33.Append(runFonts68);
            paragraphMarkRunProperties33.Append(fontSize43);
            paragraphMarkRunProperties33.Append(fontSizeComplexScript43);

            paragraphProperties33.Append(paragraphMarkRunProperties33);

            paragraph33.Append(paragraphProperties33);

            tableCell30.Append(tableCellProperties30);
            tableCell30.Append(paragraph33);

            TableCell tableCell31 = new TableCell();

            TableCellProperties tableCellProperties31 = new TableCellProperties();
            TableCellWidth tableCellWidth31 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge19 = new VerticalMerge();

            TableCellBorders tableCellBorders31 = new TableCellBorders();
            LeftBorder leftBorder32 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder18 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder32 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders31.Append(leftBorder32);
            tableCellBorders31.Append(bottomBorder18);
            tableCellBorders31.Append(rightBorder32);
            Shading shading31 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment31 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties31.Append(tableCellWidth31);
            tableCellProperties31.Append(verticalMerge19);
            tableCellProperties31.Append(tableCellBorders31);
            tableCellProperties31.Append(shading31);
            tableCellProperties31.Append(tableCellVerticalAlignment31);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts69 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize44 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties34.Append(runFonts69);
            paragraphMarkRunProperties34.Append(fontSize44);
            paragraphMarkRunProperties34.Append(fontSizeComplexScript44);

            paragraphProperties34.Append(paragraphMarkRunProperties34);

            paragraph34.Append(paragraphProperties34);

            tableCell31.Append(tableCellProperties31);
            tableCell31.Append(paragraph34);

            TableCell tableCell32 = new TableCell();

            TableCellProperties tableCellProperties32 = new TableCellProperties();
            TableCellWidth tableCellWidth32 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders32 = new TableCellBorders();
            TopBorder topBorder23 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder33 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder19 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder33 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders32.Append(topBorder23);
            tableCellBorders32.Append(leftBorder33);
            tableCellBorders32.Append(bottomBorder19);
            tableCellBorders32.Append(rightBorder33);
            Shading shading32 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment32 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties32.Append(tableCellWidth32);
            tableCellProperties32.Append(tableCellBorders32);
            tableCellProperties32.Append(shading32);
            tableCellProperties32.Append(tableCellVerticalAlignment32);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts70 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold11 = new Bold();
            FontSize fontSize45 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties35.Append(runFonts70);
            paragraphMarkRunProperties35.Append(bold11);
            paragraphMarkRunProperties35.Append(fontSize45);
            paragraphMarkRunProperties35.Append(fontSizeComplexScript45);

            paragraphProperties35.Append(paragraphMarkRunProperties35);

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<w:smartTag w:uri=\"urn:schemas-microsoft-com:office:smarttags\" w:element=\"place\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r w:rsidRPr=\"00ED7FB2\"><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\" /><w:b /><w:sz w:val=\"20\" /><w:szCs w:val=\"20\" /></w:rPr><w:t>Lot</w:t></w:r></w:smartTag>");

            Run run36 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts71 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold12 = new Bold();
            FontSize fontSize46 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "20" };

            runProperties36.Append(runFonts71);
            runProperties36.Append(bold12);
            runProperties36.Append(fontSize46);
            runProperties36.Append(fontSizeComplexScript46);
            Text text36 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text36.Text = " #’s";

            run36.Append(runProperties36);
            run36.Append(text36);

            paragraph35.Append(paragraphProperties35);
            paragraph35.Append(openXmlUnknownElement2);
            paragraph35.Append(run36);

            tableCell32.Append(tableCellProperties32);
            tableCell32.Append(paragraph35);

            TableCell tableCell33 = new TableCell();

            TableCellProperties tableCellProperties33 = new TableCellProperties();
            TableCellWidth tableCellWidth33 = new TableCellWidth() { Width = "2880", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 3 };

            TableCellBorders tableCellBorders33 = new TableCellBorders();
            LeftBorder leftBorder34 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder34 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders33.Append(leftBorder34);
            tableCellBorders33.Append(rightBorder34);
            Shading shading33 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment33 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties33.Append(tableCellWidth33);
            tableCellProperties33.Append(gridSpan1);
            tableCellProperties33.Append(tableCellBorders33);
            tableCellProperties33.Append(shading33);
            tableCellProperties33.Append(tableCellVerticalAlignment33);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic6 = new Italic();
            FontSize fontSize47 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties36.Append(runFonts72);
            paragraphMarkRunProperties36.Append(italic6);
            paragraphMarkRunProperties36.Append(fontSize47);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript47);

            paragraphProperties36.Append(paragraphMarkRunProperties36);

            Run run37 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts73 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize48 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "20" };

            runProperties37.Append(runFonts73);
            runProperties37.Append(fontSize48);
            runProperties37.Append(fontSizeComplexScript48);
            Text text37 = new Text();
            text37.Text = labSheetA1Sheet.ControlLot;

            run37.Append(runProperties37);
            run37.Append(text37);

            paragraph36.Append(paragraphProperties36);
            paragraph36.Append(run37);

            tableCell33.Append(tableCellProperties33);
            tableCell33.Append(paragraph36);

            TableCell tableCell34 = new TableCell();

            TableCellProperties tableCellProperties34 = new TableCellProperties();
            TableCellWidth tableCellWidth34 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge20 = new VerticalMerge();

            TableCellBorders tableCellBorders34 = new TableCellBorders();
            LeftBorder leftBorder35 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder20 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder35 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders34.Append(leftBorder35);
            tableCellBorders34.Append(bottomBorder20);
            tableCellBorders34.Append(rightBorder35);
            Shading shading34 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment34 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties34.Append(tableCellWidth34);
            tableCellProperties34.Append(verticalMerge20);
            tableCellProperties34.Append(tableCellBorders34);
            tableCellProperties34.Append(shading34);
            tableCellProperties34.Append(tableCellVerticalAlignment34);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts76 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize51 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties37.Append(runFonts76);
            paragraphMarkRunProperties37.Append(fontSize51);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript51);

            paragraphProperties37.Append(paragraphMarkRunProperties37);

            paragraph37.Append(paragraphProperties37);

            tableCell34.Append(tableCellProperties34);
            tableCell34.Append(paragraph37);

            TableCell tableCell35 = new TableCell();

            TableCellProperties tableCellProperties35 = new TableCellProperties();
            TableCellWidth tableCellWidth35 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge21 = new VerticalMerge();

            TableCellBorders tableCellBorders35 = new TableCellBorders();
            LeftBorder leftBorder36 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder21 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder36 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders35.Append(leftBorder36);
            tableCellBorders35.Append(bottomBorder21);
            tableCellBorders35.Append(rightBorder36);
            Shading shading35 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment35 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties35.Append(tableCellWidth35);
            tableCellProperties35.Append(verticalMerge21);
            tableCellProperties35.Append(tableCellBorders35);
            tableCellProperties35.Append(shading35);
            tableCellProperties35.Append(tableCellVerticalAlignment35);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts77 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize52 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties38.Append(runFonts77);
            paragraphMarkRunProperties38.Append(fontSize52);
            paragraphMarkRunProperties38.Append(fontSizeComplexScript52);

            paragraphProperties38.Append(paragraphMarkRunProperties38);

            paragraph38.Append(paragraphProperties38);

            tableCell35.Append(tableCellProperties35);
            tableCell35.Append(paragraph38);

            TableCell tableCell36 = new TableCell();

            TableCellProperties tableCellProperties36 = new TableCellProperties();
            TableCellWidth tableCellWidth36 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge22 = new VerticalMerge();

            TableCellBorders tableCellBorders36 = new TableCellBorders();
            LeftBorder leftBorder37 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder22 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder37 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders36.Append(leftBorder37);
            tableCellBorders36.Append(bottomBorder22);
            tableCellBorders36.Append(rightBorder37);
            Shading shading36 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment36 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties36.Append(tableCellWidth36);
            tableCellProperties36.Append(verticalMerge22);
            tableCellProperties36.Append(tableCellBorders36);
            tableCellProperties36.Append(shading36);
            tableCellProperties36.Append(tableCellVerticalAlignment36);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            RunFonts runFonts78 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize53 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties39.Append(runFonts78);
            paragraphMarkRunProperties39.Append(fontSize53);
            paragraphMarkRunProperties39.Append(fontSizeComplexScript53);

            paragraphProperties39.Append(paragraphMarkRunProperties39);

            paragraph39.Append(paragraphProperties39);

            tableCell36.Append(tableCellProperties36);
            tableCell36.Append(paragraph39);

            tableRow3.Append(tableRowProperties2);
            tableRow3.Append(tableCell24);
            tableRow3.Append(tableCell25);
            tableRow3.Append(tableCell26);
            tableRow3.Append(tableCell27);
            tableRow3.Append(tableCell28);
            tableRow3.Append(tableCell29);
            tableRow3.Append(tableCell30);
            tableRow3.Append(tableCell31);
            tableRow3.Append(tableCell32);
            tableRow3.Append(tableCell33);
            tableRow3.Append(tableCell34);
            tableRow3.Append(tableCell35);
            tableRow3.Append(tableCell36);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00397699", RsidTableRowProperties = "00ED7FB2" };

            TableCell tableCell37 = new TableCell();

            TableCellProperties tableCellProperties37 = new TableCellProperties();
            TableCellWidth tableCellWidth37 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders37 = new TableCellBorders();
            TopBorder topBorder24 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder38 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder23 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder38 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders37.Append(topBorder24);
            tableCellBorders37.Append(leftBorder38);
            tableCellBorders37.Append(bottomBorder23);
            tableCellBorders37.Append(rightBorder38);
            Shading shading37 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment37 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties37.Append(tableCellWidth37);
            tableCellProperties37.Append(tableCellBorders37);
            tableCellProperties37.Append(shading37);
            tableCellProperties37.Append(tableCellVerticalAlignment37);

            Paragraph paragraph40 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            RunFonts runFonts79 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize54 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties40.Append(runFonts79);
            paragraphMarkRunProperties40.Append(fontSize54);
            paragraphMarkRunProperties40.Append(fontSizeComplexScript54);

            paragraphProperties40.Append(paragraphMarkRunProperties40);

            Run run40 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties40 = new RunProperties();
            RunFonts runFonts80 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize55 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "20" };

            runProperties40.Append(runFonts80);
            runProperties40.Append(fontSize55);
            runProperties40.Append(fontSizeComplexScript55);
            Text text40 = new Text();
            text40.Text = "Tide";

            run40.Append(runProperties40);
            run40.Append(text40);

            paragraph40.Append(paragraphProperties40);
            paragraph40.Append(run40);

            tableCell37.Append(tableCellProperties37);
            tableCell37.Append(paragraph40);

            TableCell tableCell38 = new TableCell();

            TableCellProperties tableCellProperties38 = new TableCellProperties();
            TableCellWidth tableCellWidth38 = new TableCellWidth() { Width = "1800", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders38 = new TableCellBorders();
            TopBorder topBorder25 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder39 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder24 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder39 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders38.Append(topBorder25);
            tableCellBorders38.Append(leftBorder39);
            tableCellBorders38.Append(bottomBorder24);
            tableCellBorders38.Append(rightBorder39);
            Shading shading38 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment38 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties38.Append(tableCellWidth38);
            tableCellProperties38.Append(tableCellBorders38);
            tableCellProperties38.Append(shading38);
            tableCellProperties38.Append(tableCellVerticalAlignment38);

            Paragraph paragraph41 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            RunFonts runFonts81 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize56 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties41.Append(runFonts81);
            paragraphMarkRunProperties41.Append(fontSize56);
            paragraphMarkRunProperties41.Append(fontSizeComplexScript56);

            paragraphProperties41.Append(justification11);
            paragraphProperties41.Append(paragraphMarkRunProperties41);

            Run run41 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts82 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize57 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "20" };

            //string ColorValue = "000000";
            //if (!AllowableTideTextList.Contains(csspWQInputToolForm.a1Sheet.Tides))
            //{
            //    ColorValue = "FF0000";
            //}

            //Color color423 = new Color() { Val = ColorValue };
            //runProperties41.Append(color423);
            runProperties41.Append(runFonts82);
            runProperties41.Append(fontSize57);
            runProperties41.Append(fontSizeComplexScript57);
            Text text41 = new Text();
            text41.Text = labSheetA1Sheet.Tides;

            run41.Append(runProperties41);
            run41.Append(text41);

            paragraph41.Append(paragraphProperties41);
            paragraph41.Append(run41);

            tableCell38.Append(tableCellProperties38);
            tableCell38.Append(paragraph41);

            TableCell tableCell39 = new TableCell();

            TableCellProperties tableCellProperties39 = new TableCellProperties();
            TableCellWidth tableCellWidth39 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders39 = new TableCellBorders();
            TopBorder topBorder26 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder40 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder25 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder40 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders39.Append(topBorder26);
            tableCellBorders39.Append(leftBorder40);
            tableCellBorders39.Append(bottomBorder25);
            tableCellBorders39.Append(rightBorder40);
            Shading shading39 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment39 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties39.Append(tableCellWidth39);
            tableCellProperties39.Append(tableCellBorders39);
            tableCellProperties39.Append(shading39);
            tableCellProperties39.Append(tableCellVerticalAlignment39);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            RunFonts runFonts85 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize60 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties42.Append(runFonts85);
            paragraphMarkRunProperties42.Append(fontSize60);
            paragraphMarkRunProperties42.Append(fontSizeComplexScript60);

            paragraphProperties42.Append(paragraphMarkRunProperties42);

            paragraph42.Append(paragraphProperties42);

            tableCell39.Append(tableCellProperties39);
            tableCell39.Append(paragraph42);

            TableCell tableCell40 = new TableCell();

            TableCellProperties tableCellProperties40 = new TableCellProperties();
            TableCellWidth tableCellWidth40 = new TableCellWidth() { Width = "2284", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders40 = new TableCellBorders();
            TopBorder topBorder27 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder41 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder26 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder41 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders40.Append(topBorder27);
            tableCellBorders40.Append(leftBorder41);
            tableCellBorders40.Append(bottomBorder26);
            tableCellBorders40.Append(rightBorder41);
            Shading shading40 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment40 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties40.Append(tableCellWidth40);
            tableCellProperties40.Append(tableCellBorders40);
            tableCellProperties40.Append(shading40);
            tableCellProperties40.Append(tableCellVerticalAlignment40);

            Paragraph paragraph43 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties43 = new ParagraphMarkRunProperties();
            RunFonts runFonts86 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize61 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties43.Append(runFonts86);
            paragraphMarkRunProperties43.Append(fontSize61);
            paragraphMarkRunProperties43.Append(fontSizeComplexScript61);

            paragraphProperties43.Append(paragraphMarkRunProperties43);

            Run run44 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties44 = new RunProperties();
            RunFonts runFonts87 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize62 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "20" };

            runProperties44.Append(runFonts87);
            runProperties44.Append(fontSize62);
            runProperties44.Append(fontSizeComplexScript62);
            Text text44 = new Text();
            text44.Text = "Incubation Finish Time";

            run44.Append(runProperties44);
            run44.Append(text44);

            paragraph43.Append(paragraphProperties43);
            paragraph43.Append(run44);

            tableCell40.Append(tableCellProperties40);
            tableCell40.Append(paragraph43);

            TableCell tableCell41 = new TableCell();

            TableCellProperties tableCellProperties41 = new TableCellProperties();
            TableCellWidth tableCellWidth41 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders41 = new TableCellBorders();
            TopBorder topBorder28 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder42 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder27 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder42 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders41.Append(topBorder28);
            tableCellBorders41.Append(leftBorder42);
            tableCellBorders41.Append(bottomBorder27);
            tableCellBorders41.Append(rightBorder42);
            Shading shading41 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment41 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties41.Append(tableCellWidth41);
            tableCellProperties41.Append(tableCellBorders41);
            tableCellProperties41.Append(shading41);
            tableCellProperties41.Append(tableCellVerticalAlignment41);

            Paragraph paragraph44 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties44 = new ParagraphMarkRunProperties();
            RunFonts runFonts88 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize63 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties44.Append(runFonts88);
            paragraphMarkRunProperties44.Append(fontSize63);
            paragraphMarkRunProperties44.Append(fontSizeComplexScript63);

            paragraphProperties44.Append(paragraphMarkRunProperties44);

            Run run45 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts89 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize64 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "20" };

            runProperties45.Append(runFonts89);
            runProperties45.Append(fontSize64);
            runProperties45.Append(fontSizeComplexScript64);
            Text text45 = new Text();
            text45.Text = labSheetA1Sheet.IncubationEndTime;

            run45.Append(runProperties45);
            run45.Append(text45);

            paragraph44.Append(paragraphProperties44);
            paragraph44.Append(run45);

            tableCell41.Append(tableCellProperties41);
            tableCell41.Append(paragraph44);

            TableCell tableCell42 = new TableCell();

            TableCellProperties tableCellProperties42 = new TableCellProperties();
            TableCellWidth tableCellWidth42 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders42 = new TableCellBorders();
            TopBorder topBorder29 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder43 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder28 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder43 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders42.Append(topBorder29);
            tableCellBorders42.Append(leftBorder43);
            tableCellBorders42.Append(bottomBorder28);
            tableCellBorders42.Append(rightBorder43);
            Shading shading42 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment42 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties42.Append(tableCellWidth42);
            tableCellProperties42.Append(tableCellBorders42);
            tableCellProperties42.Append(shading42);
            tableCellProperties42.Append(tableCellVerticalAlignment42);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties45 = new ParagraphMarkRunProperties();
            RunFonts runFonts94 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color4 = new Color() { Val = "DDDDDD" };
            FontSize fontSize69 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties45.Append(runFonts94);
            paragraphMarkRunProperties45.Append(color4);
            paragraphMarkRunProperties45.Append(fontSize69);
            paragraphMarkRunProperties45.Append(fontSizeComplexScript69);

            paragraphProperties45.Append(paragraphMarkRunProperties45);

            paragraph45.Append(paragraphProperties45);

            tableCell42.Append(tableCellProperties42);
            tableCell42.Append(paragraph45);

            TableCell tableCell43 = new TableCell();

            TableCellProperties tableCellProperties43 = new TableCellProperties();
            TableCellWidth tableCellWidth43 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders43 = new TableCellBorders();
            TopBorder topBorder30 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder44 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder29 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder44 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders43.Append(topBorder30);
            tableCellBorders43.Append(leftBorder44);
            tableCellBorders43.Append(bottomBorder29);
            tableCellBorders43.Append(rightBorder44);
            Shading shading43 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment43 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties43.Append(tableCellWidth43);
            tableCellProperties43.Append(tableCellBorders43);
            tableCellProperties43.Append(shading43);
            tableCellProperties43.Append(tableCellVerticalAlignment43);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties46 = new ParagraphProperties();
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties46 = new ParagraphMarkRunProperties();
            RunFonts runFonts95 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color5 = new Color() { Val = "C0C0C0" };
            FontSize fontSize70 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties46.Append(runFonts95);
            paragraphMarkRunProperties46.Append(color5);
            paragraphMarkRunProperties46.Append(fontSize70);
            paragraphMarkRunProperties46.Append(fontSizeComplexScript70);

            paragraphProperties46.Append(justification12);
            paragraphProperties46.Append(paragraphMarkRunProperties46);

            Run run50 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties50 = new RunProperties();
            RunFonts runFonts96 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize71 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "20" };

            runProperties50.Append(runFonts96);
            runProperties50.Append(fontSize71);
            runProperties50.Append(fontSizeComplexScript71);
            Text text50 = new Text();
            text50.Text = labSheetA1Sheet.TCField1 + (labSheetA1Sheet.TCField2 != null ? " / " + labSheetA1Sheet.TCField2 : "");

            run50.Append(runProperties50);
            run50.Append(text50);

            paragraph46.Append(paragraphProperties46);
            paragraph46.Append(run50);

            tableCell43.Append(tableCellProperties43);
            tableCell43.Append(paragraph46);

            TableCell tableCell44 = new TableCell();

            TableCellProperties tableCellProperties44 = new TableCellProperties();
            TableCellWidth tableCellWidth44 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders44 = new TableCellBorders();
            TopBorder topBorder31 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder45 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder30 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder45 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders44.Append(topBorder31);
            tableCellBorders44.Append(leftBorder45);
            tableCellBorders44.Append(bottomBorder30);
            tableCellBorders44.Append(rightBorder45);
            Shading shading44 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment44 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties44.Append(tableCellWidth44);
            tableCellProperties44.Append(tableCellBorders44);
            tableCellProperties44.Append(shading44);
            tableCellProperties44.Append(tableCellVerticalAlignment44);

            Paragraph paragraph47 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties47 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties47 = new ParagraphMarkRunProperties();
            RunFonts runFonts99 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic7 = new Italic();
            FontSize fontSize74 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties47.Append(runFonts99);
            paragraphMarkRunProperties47.Append(italic7);
            paragraphMarkRunProperties47.Append(fontSize74);
            paragraphMarkRunProperties47.Append(fontSizeComplexScript74);

            paragraphProperties47.Append(paragraphMarkRunProperties47);

            paragraph47.Append(paragraphProperties47);

            tableCell44.Append(tableCellProperties44);
            tableCell44.Append(paragraph47);

            TableCell tableCell45 = new TableCell();

            TableCellProperties tableCellProperties45 = new TableCellProperties();
            TableCellWidth tableCellWidth45 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders45 = new TableCellBorders();
            TopBorder topBorder32 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder46 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder31 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder46 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders45.Append(topBorder32);
            tableCellBorders45.Append(leftBorder46);
            tableCellBorders45.Append(bottomBorder31);
            tableCellBorders45.Append(rightBorder46);
            Shading shading45 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment45 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties45.Append(tableCellWidth45);
            tableCellProperties45.Append(tableCellBorders45);
            tableCellProperties45.Append(shading45);
            tableCellProperties45.Append(tableCellVerticalAlignment45);

            Paragraph paragraph48 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties48 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties48 = new ParagraphMarkRunProperties();
            RunFonts runFonts100 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize75 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties48.Append(runFonts100);
            paragraphMarkRunProperties48.Append(fontSize75);
            paragraphMarkRunProperties48.Append(fontSizeComplexScript75);

            paragraphProperties48.Append(paragraphMarkRunProperties48);

            Run run53 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts101 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize76 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "20" };

            runProperties53.Append(runFonts101);
            runProperties53.Append(fontSize76);
            runProperties53.Append(fontSizeComplexScript76);
            Text text53 = new Text();
            text53.Text = "35.0";

            run53.Append(runProperties53);
            run53.Append(text53);

            Run run54 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize77 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "20" };
            VerticalTextAlignment verticalTextAlignment2 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties54.Append(runFonts102);
            runProperties54.Append(fontSize77);
            runProperties54.Append(fontSizeComplexScript77);
            runProperties54.Append(verticalTextAlignment2);
            Text text54 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text54.Text = " o";

            run54.Append(runProperties54);
            run54.Append(text54);

            Run run55 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts103 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize78 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "20" };

            runProperties55.Append(runFonts103);
            runProperties55.Append(fontSize78);
            runProperties55.Append(fontSizeComplexScript78);
            Text text55 = new Text();
            text55.Text = "C";

            run55.Append(runProperties55);
            run55.Append(text55);

            paragraph48.Append(paragraphProperties48);
            paragraph48.Append(run53);
            paragraph48.Append(run54);
            paragraph48.Append(run55);

            tableCell45.Append(tableCellProperties45);
            tableCell45.Append(paragraph48);

            TableCell tableCell46 = new TableCell();

            TableCellProperties tableCellProperties46 = new TableCellProperties();
            TableCellWidth tableCellWidth46 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders46 = new TableCellBorders();
            LeftBorder leftBorder47 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder47 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders46.Append(leftBorder47);
            tableCellBorders46.Append(rightBorder47);
            Shading shading46 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment46 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties46.Append(tableCellWidth46);
            tableCellProperties46.Append(tableCellBorders46);
            tableCellProperties46.Append(shading46);
            tableCellProperties46.Append(tableCellVerticalAlignment46);

            Paragraph paragraph49 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties49 = new ParagraphProperties();
            Justification justification13 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties49 = new ParagraphMarkRunProperties();
            RunFonts runFonts104 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic8 = new Italic();
            FontSize fontSize79 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties49.Append(runFonts104);
            paragraphMarkRunProperties49.Append(italic8);
            paragraphMarkRunProperties49.Append(fontSize79);
            paragraphMarkRunProperties49.Append(fontSizeComplexScript79);

            paragraphProperties49.Append(justification13);
            paragraphProperties49.Append(paragraphMarkRunProperties49);

            Run run56 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts105 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize80 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "20" };

            runProperties56.Append(runFonts105);
            runProperties56.Append(fontSize80);
            runProperties56.Append(fontSizeComplexScript80);
            Text text56 = new Text();
            text56.Text = labSheetA1Sheet.Positive35;

            run56.Append(runProperties56);
            run56.Append(text56);

            paragraph49.Append(paragraphProperties49);
            paragraph49.Append(run56);

            tableCell46.Append(tableCellProperties46);
            tableCell46.Append(paragraph49);

            TableCell tableCell47 = new TableCell();

            TableCellProperties tableCellProperties47 = new TableCellProperties();
            TableCellWidth tableCellWidth47 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders47 = new TableCellBorders();
            LeftBorder leftBorder48 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder48 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders47.Append(leftBorder48);
            tableCellBorders47.Append(rightBorder48);
            Shading shading47 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment47 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties47.Append(tableCellWidth47);
            tableCellProperties47.Append(tableCellBorders47);
            tableCellProperties47.Append(shading47);
            tableCellProperties47.Append(tableCellVerticalAlignment47);

            Paragraph paragraph50 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            Justification justification14 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties50 = new ParagraphMarkRunProperties();
            RunFonts runFonts108 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic9 = new Italic();
            FontSize fontSize83 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties50.Append(runFonts108);
            paragraphMarkRunProperties50.Append(italic9);
            paragraphMarkRunProperties50.Append(fontSize83);
            paragraphMarkRunProperties50.Append(fontSizeComplexScript83);

            paragraphProperties50.Append(justification14);
            paragraphProperties50.Append(paragraphMarkRunProperties50);

            Run run59 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties59 = new RunProperties();
            RunFonts runFonts109 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize84 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "20" };

            runProperties59.Append(runFonts109);
            runProperties59.Append(fontSize84);
            runProperties59.Append(fontSizeComplexScript84);
            Text text59 = new Text();
            text59.Text = labSheetA1Sheet.NonTarget35;

            run59.Append(runProperties59);
            run59.Append(text59);

            paragraph50.Append(paragraphProperties50);
            paragraph50.Append(run59);

            tableCell47.Append(tableCellProperties47);
            tableCell47.Append(paragraph50);

            TableCell tableCell48 = new TableCell();

            TableCellProperties tableCellProperties48 = new TableCellProperties();
            TableCellWidth tableCellWidth48 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders48 = new TableCellBorders();
            LeftBorder leftBorder49 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder49 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders48.Append(leftBorder49);
            tableCellBorders48.Append(rightBorder49);
            Shading shading48 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment48 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties48.Append(tableCellWidth48);
            tableCellProperties48.Append(tableCellBorders48);
            tableCellProperties48.Append(shading48);
            tableCellProperties48.Append(tableCellVerticalAlignment48);

            Paragraph paragraph51 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties51 = new ParagraphProperties();
            Justification justification15 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties51 = new ParagraphMarkRunProperties();
            RunFonts runFonts114 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic10 = new Italic();
            FontSize fontSize89 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties51.Append(runFonts114);
            paragraphMarkRunProperties51.Append(italic10);
            paragraphMarkRunProperties51.Append(fontSize89);
            paragraphMarkRunProperties51.Append(fontSizeComplexScript89);

            paragraphProperties51.Append(justification15);
            paragraphProperties51.Append(paragraphMarkRunProperties51);

            Run run64 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties64 = new RunProperties();
            RunFonts runFonts115 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize90 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "20" };

            runProperties64.Append(runFonts115);
            runProperties64.Append(fontSize90);
            runProperties64.Append(fontSizeComplexScript90);
            Text text64 = new Text();
            text64.Text = labSheetA1Sheet.Negative35;

            run64.Append(runProperties64);
            run64.Append(text64);

            paragraph51.Append(paragraphProperties51);
            paragraph51.Append(run64);

            tableCell48.Append(tableCellProperties48);
            tableCell48.Append(paragraph51);

            TableCell tableCell49 = new TableCell();

            TableCellProperties tableCellProperties49 = new TableCellProperties();
            TableCellWidth tableCellWidth49 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders49 = new TableCellBorders();
            TopBorder topBorder33 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder50 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder32 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder50 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders49.Append(topBorder33);
            tableCellBorders49.Append(leftBorder50);
            tableCellBorders49.Append(bottomBorder32);
            tableCellBorders49.Append(rightBorder50);
            Shading shading49 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment49 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties49.Append(tableCellWidth49);
            tableCellProperties49.Append(tableCellBorders49);
            tableCellProperties49.Append(shading49);
            tableCellProperties49.Append(tableCellVerticalAlignment49);

            Paragraph paragraph52 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties52 = new ParagraphProperties();
            Justification justification16 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties52 = new ParagraphMarkRunProperties();
            RunFonts runFonts119 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize94 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties52.Append(runFonts119);
            paragraphMarkRunProperties52.Append(fontSize94);
            paragraphMarkRunProperties52.Append(fontSizeComplexScript94);

            paragraphProperties52.Append(justification16);
            paragraphProperties52.Append(paragraphMarkRunProperties52);

            Run run68 = new Run();

            RunProperties runProperties68 = new RunProperties();
            RunFonts runFonts120 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize95 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "20" };

            runProperties68.Append(runFonts120);
            runProperties68.Append(fontSize95);
            runProperties68.Append(fontSizeComplexScript95);
            Text text68 = new Text();
            text68.Text = labSheetA1Sheet.Blank35;

            run68.Append(runProperties68);
            run68.Append(text68);

            paragraph52.Append(paragraphProperties52);
            paragraph52.Append(run68);

            tableCell49.Append(tableCellProperties49);
            tableCell49.Append(paragraph52);

            TableCell tableCell50 = new TableCell();

            TableCellProperties tableCellProperties50 = new TableCellProperties();
            TableCellWidth tableCellWidth50 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders50 = new TableCellBorders();
            TopBorder topBorder34 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder51 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder33 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder51 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders50.Append(topBorder34);
            tableCellBorders50.Append(leftBorder51);
            tableCellBorders50.Append(bottomBorder33);
            tableCellBorders50.Append(rightBorder51);
            Shading shading50 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment50 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties50.Append(tableCellWidth50);
            tableCellProperties50.Append(tableCellBorders50);
            tableCellProperties50.Append(shading50);
            tableCellProperties50.Append(tableCellVerticalAlignment50);

            Paragraph paragraph53 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties53 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties53 = new ParagraphMarkRunProperties();
            RunFonts runFonts121 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize96 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties53.Append(runFonts121);
            paragraphMarkRunProperties53.Append(fontSize96);
            paragraphMarkRunProperties53.Append(fontSizeComplexScript96);

            paragraphProperties53.Append(paragraphMarkRunProperties53);

            Run run69 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties69 = new RunProperties();
            RunFonts runFonts122 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize97 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript97 = new FontSizeComplexScript() { Val = "20" };

            runProperties69.Append(runFonts122);
            runProperties69.Append(fontSize97);
            runProperties69.Append(fontSizeComplexScript97);
            Text text69 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text69.Text = " 1X";

            run69.Append(runProperties69);
            run69.Append(text69);

            paragraph53.Append(paragraphProperties53);
            paragraph53.Append(run69);

            tableCell50.Append(tableCellProperties50);
            tableCell50.Append(paragraph53);

            TableCell tableCell51 = new TableCell();

            TableCellProperties tableCellProperties51 = new TableCellProperties();
            TableCellWidth tableCellWidth51 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders51 = new TableCellBorders();
            TopBorder topBorder35 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder52 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder34 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder52 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders51.Append(topBorder35);
            tableCellBorders51.Append(leftBorder52);
            tableCellBorders51.Append(bottomBorder34);
            tableCellBorders51.Append(rightBorder52);
            Shading shading51 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment51 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties51.Append(tableCellWidth51);
            tableCellProperties51.Append(tableCellBorders51);
            tableCellProperties51.Append(shading51);
            tableCellProperties51.Append(tableCellVerticalAlignment51);

            Paragraph paragraph54 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties54 = new ParagraphProperties();
            Justification justification17 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties54 = new ParagraphMarkRunProperties();
            RunFonts runFonts123 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize98 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript98 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties54.Append(runFonts123);
            paragraphMarkRunProperties54.Append(fontSize98);
            paragraphMarkRunProperties54.Append(fontSizeComplexScript98);

            paragraphProperties54.Append(justification17);
            paragraphProperties54.Append(paragraphMarkRunProperties54);

            Run run70 = new Run();

            RunProperties runProperties70 = new RunProperties();
            RunFonts runFonts124 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize99 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript99 = new FontSizeComplexScript() { Val = "20" };

            runProperties70.Append(runFonts124);
            runProperties70.Append(fontSize99);
            runProperties70.Append(fontSizeComplexScript99);
            Text text70 = new Text();
            text70.Text = labSheetA1Sheet.Lot35;

            run70.Append(runProperties70);
            run70.Append(text70);

            paragraph54.Append(paragraphProperties54);
            paragraph54.Append(run70);

            tableCell51.Append(tableCellProperties51);
            tableCell51.Append(paragraph54);

            tableRow4.Append(tableCell37);
            tableRow4.Append(tableCell38);
            tableRow4.Append(tableCell39);
            tableRow4.Append(tableCell40);
            tableRow4.Append(tableCell41);
            tableRow4.Append(tableCell42);
            tableRow4.Append(tableCell43);
            tableRow4.Append(tableCell44);
            tableRow4.Append(tableCell45);
            tableRow4.Append(tableCell46);
            tableRow4.Append(tableCell47);
            tableRow4.Append(tableCell48);
            tableRow4.Append(tableCell49);
            tableRow4.Append(tableCell50);
            tableRow4.Append(tableCell51);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00397699", RsidTableRowProperties = "00D53280" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)224U };

            tableRowProperties3.Append(tableRowHeight3);

            TableCell tableCell52 = new TableCell();

            TableCellProperties tableCellProperties52 = new TableCellProperties();
            TableCellWidth tableCellWidth52 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders52 = new TableCellBorders();
            TopBorder topBorder36 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder53 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder35 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder53 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders52.Append(topBorder36);
            tableCellBorders52.Append(leftBorder53);
            tableCellBorders52.Append(bottomBorder35);
            tableCellBorders52.Append(rightBorder53);
            Shading shading52 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment52 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties52.Append(tableCellWidth52);
            tableCellProperties52.Append(tableCellBorders52);
            tableCellProperties52.Append(shading52);
            tableCellProperties52.Append(tableCellVerticalAlignment52);

            Paragraph paragraph55 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties55 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties55 = new ParagraphMarkRunProperties();
            RunFonts runFonts127 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize102 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript102 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties55.Append(runFonts127);
            paragraphMarkRunProperties55.Append(fontSize102);
            paragraphMarkRunProperties55.Append(fontSizeComplexScript102);

            paragraphProperties55.Append(paragraphMarkRunProperties55);

            Run run73 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties73 = new RunProperties();
            RunFonts runFonts128 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize103 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript103 = new FontSizeComplexScript() { Val = "16" };

            runProperties73.Append(runFonts128);
            runProperties73.Append(fontSize103);
            runProperties73.Append(fontSizeComplexScript103);
            Text text73 = new Text();
            text73.Text = "Sample Crew";

            run73.Append(runProperties73);
            run73.Append(text73);

            paragraph55.Append(paragraphProperties55);
            paragraph55.Append(run73);

            tableCell52.Append(tableCellProperties52);
            tableCell52.Append(paragraph55);

            TableCell tableCell53 = new TableCell();

            TableCellProperties tableCellProperties53 = new TableCellProperties();
            TableCellWidth tableCellWidth53 = new TableCellWidth() { Width = "1800", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders53 = new TableCellBorders();
            TopBorder topBorder37 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder54 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder36 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder54 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders53.Append(topBorder37);
            tableCellBorders53.Append(leftBorder54);
            tableCellBorders53.Append(bottomBorder36);
            tableCellBorders53.Append(rightBorder54);
            Shading shading53 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment53 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties53.Append(tableCellWidth53);
            tableCellProperties53.Append(tableCellBorders53);
            tableCellProperties53.Append(shading53);
            tableCellProperties53.Append(tableCellVerticalAlignment53);

            Paragraph paragraph56 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties56 = new ParagraphProperties();
            Justification justification18 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties56 = new ParagraphMarkRunProperties();
            RunFonts runFonts129 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize104 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript104 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties56.Append(runFonts129);
            paragraphMarkRunProperties56.Append(fontSize104);
            paragraphMarkRunProperties56.Append(fontSizeComplexScript104);

            paragraphProperties56.Append(justification18);
            paragraphProperties56.Append(paragraphMarkRunProperties56);

            Run run74 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties74 = new RunProperties();
            RunFonts runFonts130 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize105 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript105 = new FontSizeComplexScript() { Val = "20" };

            runProperties74.Append(runFonts130);
            runProperties74.Append(fontSize105);
            runProperties74.Append(fontSizeComplexScript105);
            Text text74 = new Text();
            text74.Text = labSheetA1Sheet.SampleCrewInitials;

            run74.Append(runProperties74);
            run74.Append(text74);

            paragraph56.Append(paragraphProperties56);
            paragraph56.Append(run74);

            tableCell53.Append(tableCellProperties53);
            tableCell53.Append(paragraph56);

            TableCell tableCell54 = new TableCell();

            TableCellProperties tableCellProperties54 = new TableCellProperties();
            TableCellWidth tableCellWidth54 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders54 = new TableCellBorders();
            TopBorder topBorder38 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder55 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder37 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder55 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders54.Append(topBorder38);
            tableCellBorders54.Append(leftBorder55);
            tableCellBorders54.Append(bottomBorder37);
            tableCellBorders54.Append(rightBorder55);
            Shading shading54 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment54 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties54.Append(tableCellWidth54);
            tableCellProperties54.Append(tableCellBorders54);
            tableCellProperties54.Append(shading54);
            tableCellProperties54.Append(tableCellVerticalAlignment54);

            Paragraph paragraph57 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties57 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties57 = new ParagraphMarkRunProperties();
            RunFonts runFonts133 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize108 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript108 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties57.Append(runFonts133);
            paragraphMarkRunProperties57.Append(fontSize108);
            paragraphMarkRunProperties57.Append(fontSizeComplexScript108);

            paragraphProperties57.Append(paragraphMarkRunProperties57);

            paragraph57.Append(paragraphProperties57);

            tableCell54.Append(tableCellProperties54);
            tableCell54.Append(paragraph57);

            TableCell tableCell55 = new TableCell();

            TableCellProperties tableCellProperties55 = new TableCellProperties();
            TableCellWidth tableCellWidth55 = new TableCellWidth() { Width = "2284", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders55 = new TableCellBorders();
            TopBorder topBorder39 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder56 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder38 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder56 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders55.Append(topBorder39);
            tableCellBorders55.Append(leftBorder56);
            tableCellBorders55.Append(bottomBorder38);
            tableCellBorders55.Append(rightBorder56);
            Shading shading55 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment55 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties55.Append(tableCellWidth55);
            tableCellProperties55.Append(tableCellBorders55);
            tableCellProperties55.Append(shading55);
            tableCellProperties55.Append(tableCellVerticalAlignment55);

            Paragraph paragraph58 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties58 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties58 = new ParagraphMarkRunProperties();
            RunFonts runFonts134 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize109 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript109 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties58.Append(runFonts134);
            paragraphMarkRunProperties58.Append(fontSize109);
            paragraphMarkRunProperties58.Append(fontSizeComplexScript109);

            paragraphProperties58.Append(paragraphMarkRunProperties58);

            Run run77 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties77 = new RunProperties();
            RunFonts runFonts135 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize110 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript110 = new FontSizeComplexScript() { Val = "20" };

            runProperties77.Append(runFonts135);
            runProperties77.Append(fontSize110);
            runProperties77.Append(fontSizeComplexScript110);
            Text text77 = new Text();
            text77.Text = "Water";

            run77.Append(runProperties77);
            run77.Append(text77);

            Run run78 = new Run() { RsidRunProperties = "00ED7FB2", RsidRunAddition = "00C47BCB" };

            RunProperties runProperties78 = new RunProperties();
            RunFonts runFonts136 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize111 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript111 = new FontSizeComplexScript() { Val = "20" };

            runProperties78.Append(runFonts136);
            runProperties78.Append(fontSize111);
            runProperties78.Append(fontSizeComplexScript111);
            Text text78 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text78.Text = " ";

            run78.Append(runProperties78);
            run78.Append(text78);

            Run run79 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties79 = new RunProperties();
            RunFonts runFonts137 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize112 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript112 = new FontSizeComplexScript() { Val = "20" };

            runProperties79.Append(runFonts137);
            runProperties79.Append(fontSize112);
            runProperties79.Append(fontSizeComplexScript112);
            Text text79 = new Text();
            text79.Text = "bath #";

            run79.Append(runProperties79);
            run79.Append(text79);

            paragraph58.Append(paragraphProperties58);
            paragraph58.Append(run77);
            paragraph58.Append(run78);
            paragraph58.Append(run79);

            tableCell55.Append(tableCellProperties55);
            tableCell55.Append(paragraph58);

            TableCell tableCell56 = new TableCell();

            TableCellProperties tableCellProperties56 = new TableCellProperties();
            TableCellWidth tableCellWidth56 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders56 = new TableCellBorders();
            TopBorder topBorder40 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder57 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder39 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder57 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders56.Append(topBorder40);
            tableCellBorders56.Append(leftBorder57);
            tableCellBorders56.Append(bottomBorder39);
            tableCellBorders56.Append(rightBorder57);
            Shading shading56 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment56 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties56.Append(tableCellWidth56);
            tableCellProperties56.Append(tableCellBorders56);
            tableCellProperties56.Append(shading56);
            tableCellProperties56.Append(tableCellVerticalAlignment56);

            Paragraph paragraph59 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties59 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties59 = new ParagraphMarkRunProperties();
            RunFonts runFonts138 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize113 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript113 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties59.Append(runFonts138);
            paragraphMarkRunProperties59.Append(fontSize113);
            paragraphMarkRunProperties59.Append(fontSizeComplexScript113);

            paragraphProperties59.Append(paragraphMarkRunProperties59);

            Run run80 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties80 = new RunProperties();
            RunFonts runFonts139 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize114 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript114 = new FontSizeComplexScript() { Val = "20" };

            runProperties80.Append(runFonts139);
            runProperties80.Append(fontSize114);
            runProperties80.Append(fontSizeComplexScript114);
            Text text80 = new Text();
            text80.Text = labSheetA1Sheet.WaterBath;

            run80.Append(runProperties80);
            run80.Append(text80);

            paragraph59.Append(paragraphProperties59);
            paragraph59.Append(run80);

            tableCell56.Append(tableCellProperties56);
            tableCell56.Append(paragraph59);

            TableCell tableCell57 = new TableCell();

            TableCellProperties tableCellProperties57 = new TableCellProperties();
            TableCellWidth tableCellWidth57 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders57 = new TableCellBorders();
            TopBorder topBorder41 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder58 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder40 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder58 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders57.Append(topBorder41);
            tableCellBorders57.Append(leftBorder58);
            tableCellBorders57.Append(bottomBorder40);
            tableCellBorders57.Append(rightBorder58);
            Shading shading57 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment57 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties57.Append(tableCellWidth57);
            tableCellProperties57.Append(tableCellBorders57);
            tableCellProperties57.Append(shading57);
            tableCellProperties57.Append(tableCellVerticalAlignment57);

            Paragraph paragraph60 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties60 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties60 = new ParagraphMarkRunProperties();
            RunFonts runFonts142 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color6 = new Color() { Val = "DDDDDD" };
            FontSize fontSize117 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript117 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties60.Append(runFonts142);
            paragraphMarkRunProperties60.Append(color6);
            paragraphMarkRunProperties60.Append(fontSize117);
            paragraphMarkRunProperties60.Append(fontSizeComplexScript117);

            paragraphProperties60.Append(paragraphMarkRunProperties60);

            paragraph60.Append(paragraphProperties60);

            tableCell57.Append(tableCellProperties57);
            tableCell57.Append(paragraph60);

            TableCell tableCell58 = new TableCell();

            TableCellProperties tableCellProperties58 = new TableCellProperties();
            TableCellWidth tableCellWidth58 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders58 = new TableCellBorders();
            TopBorder topBorder42 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder59 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder41 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder59 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders58.Append(topBorder42);
            tableCellBorders58.Append(leftBorder59);
            tableCellBorders58.Append(bottomBorder41);
            tableCellBorders58.Append(rightBorder59);
            Shading shading58 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment58 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties58.Append(tableCellWidth58);
            tableCellProperties58.Append(tableCellBorders58);
            tableCellProperties58.Append(shading58);
            tableCellProperties58.Append(tableCellVerticalAlignment58);

            Paragraph paragraph61 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties61 = new ParagraphProperties();
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties61 = new ParagraphMarkRunProperties();
            RunFonts runFonts143 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color7 = new Color() { Val = "C0C0C0" };
            FontSize fontSize118 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript118 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties61.Append(runFonts143);
            paragraphMarkRunProperties61.Append(color7);
            paragraphMarkRunProperties61.Append(fontSize118);
            paragraphMarkRunProperties61.Append(fontSizeComplexScript118);

            paragraphProperties61.Append(justification19);
            paragraphProperties61.Append(paragraphMarkRunProperties61);

            Run run83 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties83 = new RunProperties();
            RunFonts runFonts144 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize119 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript119 = new FontSizeComplexScript() { Val = "20" };

            runProperties83.Append(runFonts144);
            runProperties83.Append(fontSize119);
            runProperties83.Append(fontSizeComplexScript119);
            Text text83 = new Text();
            text83.Text = labSheetA1Sheet.TCLab1 + (labSheetA1Sheet.TCLab2 != null ? " / " + labSheetA1Sheet.TCLab2 : "");

            run83.Append(runProperties83);
            run83.Append(text83);

            paragraph61.Append(paragraphProperties61);
            paragraph61.Append(run83);

            tableCell58.Append(tableCellProperties58);
            tableCell58.Append(paragraph61);

            TableCell tableCell59 = new TableCell();

            TableCellProperties tableCellProperties59 = new TableCellProperties();
            TableCellWidth tableCellWidth59 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders59 = new TableCellBorders();
            TopBorder topBorder43 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder60 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder42 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder60 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders59.Append(topBorder43);
            tableCellBorders59.Append(leftBorder60);
            tableCellBorders59.Append(bottomBorder42);
            tableCellBorders59.Append(rightBorder60);
            Shading shading59 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment59 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties59.Append(tableCellWidth59);
            tableCellProperties59.Append(tableCellBorders59);
            tableCellProperties59.Append(shading59);
            tableCellProperties59.Append(tableCellVerticalAlignment59);

            Paragraph paragraph62 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties62 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties62 = new ParagraphMarkRunProperties();
            RunFonts runFonts147 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic11 = new Italic();
            FontSize fontSize122 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript122 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties62.Append(runFonts147);
            paragraphMarkRunProperties62.Append(italic11);
            paragraphMarkRunProperties62.Append(fontSize122);
            paragraphMarkRunProperties62.Append(fontSizeComplexScript122);

            paragraphProperties62.Append(paragraphMarkRunProperties62);

            paragraph62.Append(paragraphProperties62);

            tableCell59.Append(tableCellProperties59);
            tableCell59.Append(paragraph62);

            TableCell tableCell60 = new TableCell();

            TableCellProperties tableCellProperties60 = new TableCellProperties();
            TableCellWidth tableCellWidth60 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders60 = new TableCellBorders();
            TopBorder topBorder44 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder61 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder43 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder61 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders60.Append(topBorder44);
            tableCellBorders60.Append(leftBorder61);
            tableCellBorders60.Append(bottomBorder43);
            tableCellBorders60.Append(rightBorder61);
            Shading shading60 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment60 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties60.Append(tableCellWidth60);
            tableCellProperties60.Append(tableCellBorders60);
            tableCellProperties60.Append(shading60);
            tableCellProperties60.Append(tableCellVerticalAlignment60);

            Paragraph paragraph63 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties63 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties63 = new ParagraphMarkRunProperties();
            RunFonts runFonts148 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize123 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript123 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties63.Append(runFonts148);
            paragraphMarkRunProperties63.Append(fontSize123);
            paragraphMarkRunProperties63.Append(fontSizeComplexScript123);

            paragraphProperties63.Append(paragraphMarkRunProperties63);

            Run run86 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties86 = new RunProperties();
            RunFonts runFonts149 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize124 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript124 = new FontSizeComplexScript() { Val = "20" };

            runProperties86.Append(runFonts149);
            runProperties86.Append(fontSize124);
            runProperties86.Append(fontSizeComplexScript124);
            Text text86 = new Text();
            text86.Text = "44.5";

            run86.Append(runProperties86);
            run86.Append(text86);

            Run run87 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties87 = new RunProperties();
            RunFonts runFonts150 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize125 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript125 = new FontSizeComplexScript() { Val = "20" };
            VerticalTextAlignment verticalTextAlignment3 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties87.Append(runFonts150);
            runProperties87.Append(fontSize125);
            runProperties87.Append(fontSizeComplexScript125);
            runProperties87.Append(verticalTextAlignment3);
            Text text87 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text87.Text = " o";

            run87.Append(runProperties87);
            run87.Append(text87);

            Run run88 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties88 = new RunProperties();
            RunFonts runFonts151 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize126 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript126 = new FontSizeComplexScript() { Val = "20" };

            runProperties88.Append(runFonts151);
            runProperties88.Append(fontSize126);
            runProperties88.Append(fontSizeComplexScript126);
            Text text88 = new Text();
            text88.Text = "C";

            run88.Append(runProperties88);
            run88.Append(text88);

            paragraph63.Append(paragraphProperties63);
            paragraph63.Append(run86);
            paragraph63.Append(run87);
            paragraph63.Append(run88);

            tableCell60.Append(tableCellProperties60);
            tableCell60.Append(paragraph63);

            TableCell tableCell61 = new TableCell();

            TableCellProperties tableCellProperties61 = new TableCellProperties();
            TableCellWidth tableCellWidth61 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders61 = new TableCellBorders();
            LeftBorder leftBorder62 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder44 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder62 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders61.Append(leftBorder62);
            tableCellBorders61.Append(bottomBorder44);
            tableCellBorders61.Append(rightBorder62);
            Shading shading61 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment61 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties61.Append(tableCellWidth61);
            tableCellProperties61.Append(tableCellBorders61);
            tableCellProperties61.Append(shading61);
            tableCellProperties61.Append(tableCellVerticalAlignment61);

            Paragraph paragraph64 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties64 = new ParagraphProperties();
            Justification justification20 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties64 = new ParagraphMarkRunProperties();
            RunFonts runFonts152 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic12 = new Italic();
            FontSize fontSize127 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript127 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties64.Append(runFonts152);
            paragraphMarkRunProperties64.Append(italic12);
            paragraphMarkRunProperties64.Append(fontSize127);
            paragraphMarkRunProperties64.Append(fontSizeComplexScript127);

            paragraphProperties64.Append(justification20);
            paragraphProperties64.Append(paragraphMarkRunProperties64);

            Run run89 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties89 = new RunProperties();
            RunFonts runFonts153 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize128 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript128 = new FontSizeComplexScript() { Val = "20" };

            runProperties89.Append(runFonts153);
            runProperties89.Append(fontSize128);
            runProperties89.Append(fontSizeComplexScript128);
            Text text89 = new Text();
            text89.Text = labSheetA1Sheet.Positive44_5;

            run89.Append(runProperties89);
            run89.Append(text89);
            paragraph64.Append(paragraphProperties64);
            paragraph64.Append(run89);

            tableCell61.Append(tableCellProperties61);
            tableCell61.Append(paragraph64);

            TableCell tableCell62 = new TableCell();

            TableCellProperties tableCellProperties62 = new TableCellProperties();
            TableCellWidth tableCellWidth62 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders62 = new TableCellBorders();
            LeftBorder leftBorder63 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder45 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder63 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders62.Append(leftBorder63);
            tableCellBorders62.Append(bottomBorder45);
            tableCellBorders62.Append(rightBorder63);
            Shading shading62 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment62 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties62.Append(tableCellWidth62);
            tableCellProperties62.Append(tableCellBorders62);
            tableCellProperties62.Append(shading62);
            tableCellProperties62.Append(tableCellVerticalAlignment62);

            Paragraph paragraph65 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties65 = new ParagraphProperties();
            Justification justification21 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties65 = new ParagraphMarkRunProperties();
            RunFonts runFonts157 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic13 = new Italic();
            FontSize fontSize132 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript132 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties65.Append(runFonts157);
            paragraphMarkRunProperties65.Append(italic13);
            paragraphMarkRunProperties65.Append(fontSize132);
            paragraphMarkRunProperties65.Append(fontSizeComplexScript132);

            paragraphProperties65.Append(justification21);
            paragraphProperties65.Append(paragraphMarkRunProperties65);

            Run run93 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties93 = new RunProperties();
            RunFonts runFonts158 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize133 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript133 = new FontSizeComplexScript() { Val = "20" };

            runProperties93.Append(runFonts158);
            runProperties93.Append(fontSize133);
            runProperties93.Append(fontSizeComplexScript133);
            Text text93 = new Text();
            text93.Text = labSheetA1Sheet.NonTarget44_5;

            run93.Append(runProperties93);
            run93.Append(text93);

            paragraph65.Append(paragraphProperties65);
            paragraph65.Append(run93);

            tableCell62.Append(tableCellProperties62);
            tableCell62.Append(paragraph65);

            TableCell tableCell63 = new TableCell();

            TableCellProperties tableCellProperties63 = new TableCellProperties();
            TableCellWidth tableCellWidth63 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders63 = new TableCellBorders();
            LeftBorder leftBorder64 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder46 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder64 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders63.Append(leftBorder64);
            tableCellBorders63.Append(bottomBorder46);
            tableCellBorders63.Append(rightBorder64);
            Shading shading63 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment63 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties63.Append(tableCellWidth63);
            tableCellProperties63.Append(tableCellBorders63);
            tableCellProperties63.Append(shading63);
            tableCellProperties63.Append(tableCellVerticalAlignment63);

            Paragraph paragraph66 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties66 = new ParagraphProperties();
            Justification justification22 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties66 = new ParagraphMarkRunProperties();
            RunFonts runFonts162 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic14 = new Italic();
            FontSize fontSize137 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript137 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties66.Append(runFonts162);
            paragraphMarkRunProperties66.Append(italic14);
            paragraphMarkRunProperties66.Append(fontSize137);
            paragraphMarkRunProperties66.Append(fontSizeComplexScript137);

            paragraphProperties66.Append(justification22);
            paragraphProperties66.Append(paragraphMarkRunProperties66);

            Run run97 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties97 = new RunProperties();
            RunFonts runFonts163 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize138 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript138 = new FontSizeComplexScript() { Val = "20" };

            runProperties97.Append(runFonts163);
            runProperties97.Append(fontSize138);
            runProperties97.Append(fontSizeComplexScript138);
            Text text97 = new Text();
            text97.Text = labSheetA1Sheet.Negative44_5;

            run97.Append(runProperties97);
            run97.Append(text97);

            Run run98 = new Run();

            paragraph66.Append(paragraphProperties66);
            paragraph66.Append(run97);

            tableCell63.Append(tableCellProperties63);
            tableCell63.Append(paragraph66);

            TableCell tableCell64 = new TableCell();

            TableCellProperties tableCellProperties64 = new TableCellProperties();
            TableCellWidth tableCellWidth64 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders64 = new TableCellBorders();
            TopBorder topBorder45 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder65 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder47 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder65 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders64.Append(topBorder45);
            tableCellBorders64.Append(leftBorder65);
            tableCellBorders64.Append(bottomBorder47);
            tableCellBorders64.Append(rightBorder65);
            Shading shading64 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment64 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties64.Append(tableCellWidth64);
            tableCellProperties64.Append(tableCellBorders64);
            tableCellProperties64.Append(shading64);
            tableCellProperties64.Append(tableCellVerticalAlignment64);

            Paragraph paragraph67 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties67 = new ParagraphProperties();
            Justification justification23 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties67 = new ParagraphMarkRunProperties();
            RunFonts runFonts166 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize141 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript141 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties67.Append(runFonts166);
            paragraphMarkRunProperties67.Append(fontSize141);
            paragraphMarkRunProperties67.Append(fontSizeComplexScript141);

            paragraphProperties67.Append(justification23);
            paragraphProperties67.Append(paragraphMarkRunProperties67);

            Run run100 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties100 = new RunProperties();
            RunFonts runFonts167 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize142 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript142 = new FontSizeComplexScript() { Val = "20" };

            runProperties100.Append(runFonts167);
            runProperties100.Append(fontSize142);
            runProperties100.Append(fontSizeComplexScript142);
            Text text100 = new Text();
            text100.Text = "N/A";

            run100.Append(runProperties100);
            run100.Append(text100);

            paragraph67.Append(paragraphProperties67);
            paragraph67.Append(run100);

            tableCell64.Append(tableCellProperties64);
            tableCell64.Append(paragraph67);

            TableCell tableCell65 = new TableCell();

            TableCellProperties tableCellProperties65 = new TableCellProperties();
            TableCellWidth tableCellWidth65 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders65 = new TableCellBorders();
            TopBorder topBorder46 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder66 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder48 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder66 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders65.Append(topBorder46);
            tableCellBorders65.Append(leftBorder66);
            tableCellBorders65.Append(bottomBorder48);
            tableCellBorders65.Append(rightBorder66);
            Shading shading65 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment65 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties65.Append(tableCellWidth65);
            tableCellProperties65.Append(tableCellBorders65);
            tableCellProperties65.Append(shading65);
            tableCellProperties65.Append(tableCellVerticalAlignment65);

            Paragraph paragraph68 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties68 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties68 = new ParagraphMarkRunProperties();
            RunFonts runFonts168 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize143 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript143 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties68.Append(runFonts168);
            paragraphMarkRunProperties68.Append(fontSize143);
            paragraphMarkRunProperties68.Append(fontSizeComplexScript143);

            paragraphProperties68.Append(paragraphMarkRunProperties68);

            Run run101 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties101 = new RunProperties();
            RunFonts runFonts169 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize144 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript144 = new FontSizeComplexScript() { Val = "20" };

            runProperties101.Append(runFonts169);
            runProperties101.Append(fontSize144);
            runProperties101.Append(fontSizeComplexScript144);
            Text text101 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text101.Text = " 2X";

            run101.Append(runProperties101);
            run101.Append(text101);

            paragraph68.Append(paragraphProperties68);
            paragraph68.Append(run101);

            tableCell65.Append(tableCellProperties65);
            tableCell65.Append(paragraph68);

            TableCell tableCell66 = new TableCell();

            TableCellProperties tableCellProperties66 = new TableCellProperties();
            TableCellWidth tableCellWidth66 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders66 = new TableCellBorders();
            TopBorder topBorder47 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder67 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder49 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder67 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders66.Append(topBorder47);
            tableCellBorders66.Append(leftBorder67);
            tableCellBorders66.Append(bottomBorder49);
            tableCellBorders66.Append(rightBorder67);
            Shading shading66 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment66 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties66.Append(tableCellWidth66);
            tableCellProperties66.Append(tableCellBorders66);
            tableCellProperties66.Append(shading66);
            tableCellProperties66.Append(tableCellVerticalAlignment66);

            Paragraph paragraph69 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "0072154A" };

            ParagraphProperties paragraphProperties69 = new ParagraphProperties();
            Justification justification24 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties69 = new ParagraphMarkRunProperties();
            RunFonts runFonts170 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize145 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript145 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties69.Append(runFonts170);
            paragraphMarkRunProperties69.Append(fontSize145);
            paragraphMarkRunProperties69.Append(fontSizeComplexScript145);

            paragraphProperties69.Append(justification24);
            paragraphProperties69.Append(paragraphMarkRunProperties69);

            Run run102 = new Run();

            RunProperties runProperties102 = new RunProperties();
            RunFonts runFonts171 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize146 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript146 = new FontSizeComplexScript() { Val = "20" };

            runProperties102.Append(runFonts171);
            runProperties102.Append(fontSize146);
            runProperties102.Append(fontSizeComplexScript146);
            Text text102 = new Text();
            text102.Text = labSheetA1Sheet.Lot44_5;

            run102.Append(runProperties102);
            run102.Append(text102);

            paragraph69.Append(paragraphProperties69);
            paragraph69.Append(run102);

            tableCell66.Append(tableCellProperties66);
            tableCell66.Append(paragraph69);

            tableRow5.Append(tableRowProperties3);
            tableRow5.Append(tableCell52);
            tableRow5.Append(tableCell53);
            tableRow5.Append(tableCell54);
            tableRow5.Append(tableCell55);
            tableRow5.Append(tableCell56);
            tableRow5.Append(tableCell57);
            tableRow5.Append(tableCell58);
            tableRow5.Append(tableCell59);
            tableRow5.Append(tableCell60);
            tableRow5.Append(tableCell61);
            tableRow5.Append(tableCell62);
            tableRow5.Append(tableCell63);
            tableRow5.Append(tableCell64);
            tableRow5.Append(tableCell65);
            tableRow5.Append(tableCell66);

            table2.Append(tableProperties2);
            table2.Append(tableGrid2);
            table2.Append(tableRow2);
            table2.Append(tableRow3);
            table2.Append(tableRow4);
            table2.Append(tableRow5);

            Paragraph paragraph70 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00ED7FB2", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00D53280" };

            ParagraphProperties paragraphProperties70 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties70 = new ParagraphMarkRunProperties();
            Vanish vanish1 = new Vanish();

            paragraphMarkRunProperties70.Append(vanish1);

            paragraphProperties70.Append(paragraphMarkRunProperties70);

            Run run105 = new Run() { RsidRunProperties = "009A6852" };

            RunProperties runProperties105 = new RunProperties();
            RunFonts runFonts174 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize149 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript149 = new FontSizeComplexScript() { Val = "22" };

            runProperties105.Append(runFonts174);
            runProperties105.Append(fontSize149);
            runProperties105.Append(fontSizeComplexScript149);
            Text text105 = new Text();
            text105.Text = "RECORDED TEMPERATURE IS ACTUAL READING + CORRECTION FACTOR";

            run105.Append(runProperties105);
            run105.Append(text105);

            paragraph70.Append(paragraphProperties70);
            paragraph70.Append(run105);

            Table table3 = new Table();

            TableProperties tableProperties3 = new TableProperties();
            TablePositionProperties tablePositionProperties1 = new TablePositionProperties() { LeftFromText = 180, RightFromText = 180, VerticalAnchor = VerticalAnchorValues.Text, HorizontalAnchor = HorizontalAnchorValues.Margin, TablePositionX = -144, TablePositionY = 138 };
            TableWidth tableWidth3 = new TableWidth() { Width = "15048", Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders2 = new TableBorders();
            TopBorder topBorder48 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder68 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder50 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder68 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders2.Append(topBorder48);
            tableBorders2.Append(leftBorder68);
            tableBorders2.Append(bottomBorder50);
            tableBorders2.Append(rightBorder68);
            tableBorders2.Append(insideHorizontalBorder2);
            tableBorders2.Append(insideVerticalBorder2);
            TableLayout tableLayout2 = new TableLayout() { Type = TableLayoutValues.Fixed };
            TableLook tableLook3 = new TableLook() { Val = "01E0" };

            tableProperties3.Append(tablePositionProperties1);
            tableProperties3.Append(tableWidth3);
            tableProperties3.Append(tableBorders2);
            tableProperties3.Append(tableLayout2);
            tableProperties3.Append(tableLook3);

            TableGrid tableGrid3 = new TableGrid();
            GridColumn gridColumn24 = new GridColumn() { Width = "756" };
            GridColumn gridColumn25 = new GridColumn() { Width = "864" };
            GridColumn gridColumn26 = new GridColumn() { Width = "1260" };
            GridColumn gridColumn27 = new GridColumn() { Width = "540" };
            GridColumn gridColumn28 = new GridColumn() { Width = "540" };
            GridColumn gridColumn29 = new GridColumn() { Width = "540" };
            GridColumn gridColumn30 = new GridColumn() { Width = "810" };
            GridColumn gridColumn31 = new GridColumn() { Width = "810" };
            GridColumn gridColumn32 = new GridColumn() { Width = "1024" };
            GridColumn gridColumn33 = new GridColumn() { Width = "236" };
            GridColumn gridColumn34 = new GridColumn() { Width = "1008" };
            GridColumn gridColumn35 = new GridColumn() { Width = "900" };
            GridColumn gridColumn36 = new GridColumn() { Width = "1260" };
            GridColumn gridColumn37 = new GridColumn() { Width = "540" };
            GridColumn gridColumn38 = new GridColumn() { Width = "720" };
            GridColumn gridColumn39 = new GridColumn() { Width = "540" };
            GridColumn gridColumn40 = new GridColumn() { Width = "900" };
            GridColumn gridColumn41 = new GridColumn() { Width = "228" };
            GridColumn gridColumn42 = new GridColumn() { Width = "492" };
            GridColumn gridColumn43 = new GridColumn() { Width = "1080" };

            tableGrid3.Append(gridColumn24);
            tableGrid3.Append(gridColumn25);
            tableGrid3.Append(gridColumn26);
            tableGrid3.Append(gridColumn27);
            tableGrid3.Append(gridColumn28);
            tableGrid3.Append(gridColumn29);
            tableGrid3.Append(gridColumn30);
            tableGrid3.Append(gridColumn31);
            tableGrid3.Append(gridColumn32);
            tableGrid3.Append(gridColumn33);
            tableGrid3.Append(gridColumn34);
            tableGrid3.Append(gridColumn35);
            tableGrid3.Append(gridColumn36);
            tableGrid3.Append(gridColumn37);
            tableGrid3.Append(gridColumn38);
            tableGrid3.Append(gridColumn39);
            tableGrid3.Append(gridColumn40);
            tableGrid3.Append(gridColumn41);
            tableGrid3.Append(gridColumn42);
            tableGrid3.Append(gridColumn43);

            TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "0077471C", RsidTableRowProperties = "00ED7FB2" };

            TableCell tableCell67 = new TableCell();

            TableCellProperties tableCellProperties67 = new TableCellProperties();
            TableCellWidth tableCellWidth67 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders67 = new TableCellBorders();
            TopBorder topBorder49 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder69 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder51 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder69 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders67.Append(topBorder49);
            tableCellBorders67.Append(leftBorder69);
            tableCellBorders67.Append(bottomBorder51);
            tableCellBorders67.Append(rightBorder69);
            Shading shading67 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment67 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties67.Append(tableCellWidth67);
            tableCellProperties67.Append(tableCellBorders67);
            tableCellProperties67.Append(shading67);
            tableCellProperties67.Append(tableCellVerticalAlignment67);

            Paragraph paragraph71 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties71 = new ParagraphProperties();
            Justification justification25 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties71 = new ParagraphMarkRunProperties();
            RunFonts runFonts175 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize150 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript150 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties71.Append(runFonts175);
            paragraphMarkRunProperties71.Append(fontSize150);
            paragraphMarkRunProperties71.Append(fontSizeComplexScript150);

            paragraphProperties71.Append(justification25);
            paragraphProperties71.Append(paragraphMarkRunProperties71);

            Run run106 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties106 = new RunProperties();
            RunFonts runFonts176 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize151 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript151 = new FontSizeComplexScript() { Val = "16" };

            runProperties106.Append(runFonts176);
            runProperties106.Append(fontSize151);
            runProperties106.Append(fontSizeComplexScript151);
            Text text106 = new Text();
            text106.Text = "Station";

            run106.Append(runProperties106);
            run106.Append(text106);

            paragraph71.Append(paragraphProperties71);
            paragraph71.Append(run106);

            tableCell67.Append(tableCellProperties67);
            tableCell67.Append(paragraph71);

            TableCell tableCell68 = new TableCell();

            TableCellProperties tableCellProperties68 = new TableCellProperties();
            TableCellWidth tableCellWidth68 = new TableCellWidth() { Width = "864", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders68 = new TableCellBorders();
            TopBorder topBorder50 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder70 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder52 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder70 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders68.Append(topBorder50);
            tableCellBorders68.Append(leftBorder70);
            tableCellBorders68.Append(bottomBorder52);
            tableCellBorders68.Append(rightBorder70);
            Shading shading68 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment68 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties68.Append(tableCellWidth68);
            tableCellProperties68.Append(tableCellBorders68);
            tableCellProperties68.Append(shading68);
            tableCellProperties68.Append(tableCellVerticalAlignment68);

            Paragraph paragraph72 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties72 = new ParagraphProperties();
            Justification justification26 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties72 = new ParagraphMarkRunProperties();
            RunFonts runFonts177 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize152 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript152 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties72.Append(runFonts177);
            paragraphMarkRunProperties72.Append(fontSize152);
            paragraphMarkRunProperties72.Append(fontSizeComplexScript152);

            paragraphProperties72.Append(justification26);
            paragraphProperties72.Append(paragraphMarkRunProperties72);

            Run run107 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties107 = new RunProperties();
            RunFonts runFonts178 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize153 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript153 = new FontSizeComplexScript() { Val = "16" };

            runProperties107.Append(runFonts178);
            runProperties107.Append(fontSize153);
            runProperties107.Append(fontSizeComplexScript153);
            Text text107 = new Text();
            text107.Text = "Time";

            run107.Append(runProperties107);
            run107.Append(text107);

            paragraph72.Append(paragraphProperties72);
            paragraph72.Append(run107);

            tableCell68.Append(tableCellProperties68);
            tableCell68.Append(paragraph72);

            TableCell tableCell69 = new TableCell();

            TableCellProperties tableCellProperties69 = new TableCellProperties();
            TableCellWidth tableCellWidth69 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders69 = new TableCellBorders();
            TopBorder topBorder51 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder71 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder53 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder71 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders69.Append(topBorder51);
            tableCellBorders69.Append(leftBorder71);
            tableCellBorders69.Append(bottomBorder53);
            tableCellBorders69.Append(rightBorder71);
            Shading shading69 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment69 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties69.Append(tableCellWidth69);
            tableCellProperties69.Append(tableCellBorders69);
            tableCellProperties69.Append(shading69);
            tableCellProperties69.Append(tableCellVerticalAlignment69);

            Paragraph paragraph73 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties73 = new ParagraphProperties();
            Justification justification27 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties73 = new ParagraphMarkRunProperties();
            RunFonts runFonts179 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize154 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript154 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties73.Append(runFonts179);
            paragraphMarkRunProperties73.Append(fontSize154);
            paragraphMarkRunProperties73.Append(fontSizeComplexScript154);

            paragraphProperties73.Append(justification27);
            paragraphProperties73.Append(paragraphMarkRunProperties73);

            Run run108 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties108 = new RunProperties();
            RunFonts runFonts180 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize155 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript155 = new FontSizeComplexScript() { Val = "16" };

            runProperties108.Append(runFonts180);
            runProperties108.Append(fontSize155);
            runProperties108.Append(fontSizeComplexScript155);
            Text text108 = new Text();
            text108.Text = "MPN/ 100ml";

            run108.Append(runProperties108);
            run108.Append(text108);

            paragraph73.Append(paragraphProperties73);
            paragraph73.Append(run108);

            Paragraph paragraph74 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties74 = new ParagraphProperties();
            Justification justification28 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties74 = new ParagraphMarkRunProperties();
            RunFonts runFonts181 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize156 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript156 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties74.Append(runFonts181);
            paragraphMarkRunProperties74.Append(fontSize156);
            paragraphMarkRunProperties74.Append(fontSizeComplexScript156);

            paragraphProperties74.Append(justification28);
            paragraphProperties74.Append(paragraphMarkRunProperties74);

            Run run109 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties109 = new RunProperties();
            RunFonts runFonts182 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize157 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript157 = new FontSizeComplexScript() { Val = "16" };

            runProperties109.Append(runFonts182);
            runProperties109.Append(fontSize157);
            runProperties109.Append(fontSizeComplexScript157);
            Text text109 = new Text();
            text109.Text = "Fecal Coliform";

            run109.Append(runProperties109);
            run109.Append(text109);

            paragraph74.Append(paragraphProperties74);
            paragraph74.Append(run109);

            tableCell69.Append(tableCellProperties69);
            tableCell69.Append(paragraph73);
            tableCell69.Append(paragraph74);

            TableCell tableCell70 = new TableCell();

            TableCellProperties tableCellProperties70 = new TableCellProperties();
            TableCellWidth tableCellWidth70 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan2 = new GridSpan() { Val = 3 };

            TableCellBorders tableCellBorders70 = new TableCellBorders();
            TopBorder topBorder52 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder72 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder54 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder72 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders70.Append(topBorder52);
            tableCellBorders70.Append(leftBorder72);
            tableCellBorders70.Append(bottomBorder54);
            tableCellBorders70.Append(rightBorder72);
            Shading shading70 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment70 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties70.Append(tableCellWidth70);
            tableCellProperties70.Append(gridSpan2);
            tableCellProperties70.Append(tableCellBorders70);
            tableCellProperties70.Append(shading70);
            tableCellProperties70.Append(tableCellVerticalAlignment70);

            Paragraph paragraph75 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties75 = new ParagraphProperties();
            Justification justification29 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties75 = new ParagraphMarkRunProperties();
            RunFonts runFonts183 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize158 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript158 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties75.Append(runFonts183);
            paragraphMarkRunProperties75.Append(fontSize158);
            paragraphMarkRunProperties75.Append(fontSizeComplexScript158);

            paragraphProperties75.Append(justification29);
            paragraphProperties75.Append(paragraphMarkRunProperties75);

            Run run110 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties110 = new RunProperties();
            RunFonts runFonts184 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize159 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript159 = new FontSizeComplexScript() { Val = "16" };

            runProperties110.Append(runFonts184);
            runProperties110.Append(fontSize159);
            runProperties110.Append(fontSizeComplexScript159);
            Text text110 = new Text();
            text110.Text = "Positive Tubes";

            run110.Append(runProperties110);
            run110.Append(text110);

            paragraph75.Append(paragraphProperties75);
            paragraph75.Append(run110);

            Paragraph paragraph76 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00D41B28" };

            ParagraphProperties paragraphProperties76 = new ParagraphProperties();
            Justification justification30 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties76 = new ParagraphMarkRunProperties();
            RunFonts runFonts185 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize160 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript160 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties76.Append(runFonts185);
            paragraphMarkRunProperties76.Append(fontSize160);
            paragraphMarkRunProperties76.Append(fontSizeComplexScript160);

            paragraphProperties76.Append(justification30);
            paragraphProperties76.Append(paragraphMarkRunProperties76);

            Run run111 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties111 = new RunProperties();
            RunFonts runFonts186 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize161 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript161 = new FontSizeComplexScript() { Val = "16" };

            runProperties111.Append(runFonts186);
            runProperties111.Append(fontSize161);
            runProperties111.Append(fontSizeComplexScript161);
            Text text111 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text111.Text = "10     1.0    ";

            run111.Append(runProperties111);
            run111.Append(text111);

            Run run112 = new Run() { RsidRunProperties = "00ED7FB2", RsidRunAddition = "00B8087D" };

            RunProperties runProperties112 = new RunProperties();
            RunFonts runFonts187 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize162 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript162 = new FontSizeComplexScript() { Val = "16" };

            runProperties112.Append(runFonts187);
            runProperties112.Append(fontSize162);
            runProperties112.Append(fontSizeComplexScript162);
            Text text112 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text112.Text = " 0.1";

            run112.Append(runProperties112);
            run112.Append(text112);

            paragraph76.Append(paragraphProperties76);
            paragraph76.Append(run111);
            paragraph76.Append(run112);

            tableCell70.Append(tableCellProperties70);
            tableCell70.Append(paragraph75);
            tableCell70.Append(paragraph76);

            TableCell tableCell71 = new TableCell();

            TableCellProperties tableCellProperties71 = new TableCellProperties();
            TableCellWidth tableCellWidth71 = new TableCellWidth() { Width = "810", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders71 = new TableCellBorders();
            TopBorder topBorder53 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder73 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder55 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder73 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders71.Append(topBorder53);
            tableCellBorders71.Append(leftBorder73);
            tableCellBorders71.Append(bottomBorder55);
            tableCellBorders71.Append(rightBorder73);
            Shading shading71 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment71 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties71.Append(tableCellWidth71);
            tableCellProperties71.Append(tableCellBorders71);
            tableCellProperties71.Append(shading71);
            tableCellProperties71.Append(tableCellVerticalAlignment71);

            Paragraph paragraph77 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties77 = new ParagraphProperties();
            Justification justification31 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties77 = new ParagraphMarkRunProperties();
            RunFonts runFonts188 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize163 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript163 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties77.Append(runFonts188);
            paragraphMarkRunProperties77.Append(fontSize163);
            paragraphMarkRunProperties77.Append(fontSizeComplexScript163);

            paragraphProperties77.Append(justification31);
            paragraphProperties77.Append(paragraphMarkRunProperties77);

            Run run113 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties113 = new RunProperties();
            RunFonts runFonts189 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize164 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript164 = new FontSizeComplexScript() { Val = "16" };

            runProperties113.Append(runFonts189);
            runProperties113.Append(fontSize164);
            runProperties113.Append(fontSizeComplexScript164);
            Text text113 = new Text();
            text113.Text = "Salinity";

            run113.Append(runProperties113);
            run113.Append(text113);

            paragraph77.Append(paragraphProperties77);
            paragraph77.Append(run113);

            Paragraph paragraph78 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties78 = new ParagraphProperties();
            Justification justification32 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties78 = new ParagraphMarkRunProperties();
            RunFonts runFonts190 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize165 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript165 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties78.Append(runFonts190);
            paragraphMarkRunProperties78.Append(fontSize165);
            paragraphMarkRunProperties78.Append(fontSizeComplexScript165);

            paragraphProperties78.Append(justification32);
            paragraphProperties78.Append(paragraphMarkRunProperties78);

            Run run114 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties114 = new RunProperties();
            RunFonts runFonts191 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize166 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript166 = new FontSizeComplexScript() { Val = "16" };

            runProperties114.Append(runFonts191);
            runProperties114.Append(fontSize166);
            runProperties114.Append(fontSizeComplexScript166);
            Text text114 = new Text();
            text114.Text = "(ppt)";

            run114.Append(runProperties114);
            run114.Append(text114);

            paragraph78.Append(paragraphProperties78);
            paragraph78.Append(run114);

            tableCell71.Append(tableCellProperties71);
            tableCell71.Append(paragraph77);
            tableCell71.Append(paragraph78);

            TableCell tableCell72 = new TableCell();

            TableCellProperties tableCellProperties72 = new TableCellProperties();
            TableCellWidth tableCellWidth72 = new TableCellWidth() { Width = "810", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders72 = new TableCellBorders();
            TopBorder topBorder54 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder74 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder56 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder74 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders72.Append(topBorder54);
            tableCellBorders72.Append(leftBorder74);
            tableCellBorders72.Append(bottomBorder56);
            tableCellBorders72.Append(rightBorder74);
            Shading shading72 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment72 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties72.Append(tableCellWidth72);
            tableCellProperties72.Append(tableCellBorders72);
            tableCellProperties72.Append(shading72);
            tableCellProperties72.Append(tableCellVerticalAlignment72);

            Paragraph paragraph79 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties79 = new ParagraphProperties();
            Justification justification33 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties79 = new ParagraphMarkRunProperties();
            RunFonts runFonts192 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize167 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript167 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties79.Append(runFonts192);
            paragraphMarkRunProperties79.Append(fontSize167);
            paragraphMarkRunProperties79.Append(fontSizeComplexScript167);

            paragraphProperties79.Append(justification33);
            paragraphProperties79.Append(paragraphMarkRunProperties79);

            Run run115 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties115 = new RunProperties();
            RunFonts runFonts193 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize168 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript168 = new FontSizeComplexScript() { Val = "16" };

            runProperties115.Append(runFonts193);
            runProperties115.Append(fontSize168);
            runProperties115.Append(fontSizeComplexScript168);
            Text text115 = new Text();
            text115.Text = "Temp";

            run115.Append(runProperties115);
            run115.Append(text115);

            paragraph79.Append(paragraphProperties79);
            paragraph79.Append(run115);

            Paragraph paragraph80 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties80 = new ParagraphProperties();
            Justification justification34 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties80 = new ParagraphMarkRunProperties();
            RunFonts runFonts194 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize169 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript169 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties80.Append(runFonts194);
            paragraphMarkRunProperties80.Append(fontSize169);
            paragraphMarkRunProperties80.Append(fontSizeComplexScript169);

            paragraphProperties80.Append(justification34);
            paragraphProperties80.Append(paragraphMarkRunProperties80);

            Run run116 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties116 = new RunProperties();
            RunFonts runFonts195 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize170 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript170 = new FontSizeComplexScript() { Val = "16" };

            runProperties116.Append(runFonts195);
            runProperties116.Append(fontSize170);
            runProperties116.Append(fontSizeComplexScript170);
            Text text116 = new Text();
            text116.Text = "(";

            run116.Append(runProperties116);
            run116.Append(text116);

            Run run117 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties117 = new RunProperties();
            RunFonts runFonts196 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize171 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript171 = new FontSizeComplexScript() { Val = "16" };
            VerticalTextAlignment verticalTextAlignment4 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties117.Append(runFonts196);
            runProperties117.Append(fontSize171);
            runProperties117.Append(fontSizeComplexScript171);
            runProperties117.Append(verticalTextAlignment4);
            Text text117 = new Text();
            text117.Text = "o";

            run117.Append(runProperties117);
            run117.Append(text117);

            Run run118 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties118 = new RunProperties();
            RunFonts runFonts197 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize172 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript172 = new FontSizeComplexScript() { Val = "16" };

            runProperties118.Append(runFonts197);
            runProperties118.Append(fontSize172);
            runProperties118.Append(fontSizeComplexScript172);
            Text text118 = new Text();
            text118.Text = "C)";

            run118.Append(runProperties118);
            run118.Append(text118);

            paragraph80.Append(paragraphProperties80);
            paragraph80.Append(run116);
            paragraph80.Append(run117);
            paragraph80.Append(run118);

            tableCell72.Append(tableCellProperties72);
            tableCell72.Append(paragraph79);
            tableCell72.Append(paragraph80);

            TableCell tableCell73 = new TableCell();

            TableCellProperties tableCellProperties73 = new TableCellProperties();
            TableCellWidth tableCellWidth73 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders73 = new TableCellBorders();
            TopBorder topBorder55 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder75 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder57 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder75 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders73.Append(topBorder55);
            tableCellBorders73.Append(leftBorder75);
            tableCellBorders73.Append(bottomBorder57);
            tableCellBorders73.Append(rightBorder75);
            Shading shading73 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment73 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties73.Append(tableCellWidth73);
            tableCellProperties73.Append(tableCellBorders73);
            tableCellProperties73.Append(shading73);
            tableCellProperties73.Append(tableCellVerticalAlignment73);

            Paragraph paragraph81 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00286EB4" };

            ParagraphProperties paragraphProperties81 = new ParagraphProperties();
            Justification justification35 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties81 = new ParagraphMarkRunProperties();
            RunFonts runFonts198 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize173 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript173 = new FontSizeComplexScript() { Val = "16" };
            Highlight highlight1 = new Highlight() { Val = HighlightColorValues.Yellow };

            paragraphMarkRunProperties81.Append(runFonts198);
            paragraphMarkRunProperties81.Append(fontSize173);
            paragraphMarkRunProperties81.Append(fontSizeComplexScript173);
            paragraphMarkRunProperties81.Append(highlight1);

            paragraphProperties81.Append(justification35);
            paragraphProperties81.Append(paragraphMarkRunProperties81);

            Run run119 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties119 = new RunProperties();
            RunFonts runFonts199 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize174 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript174 = new FontSizeComplexScript() { Val = "16" };

            runProperties119.Append(runFonts199);
            runProperties119.Append(fontSize174);
            runProperties119.Append(fontSizeComplexScript174);
            Text text119 = new Text();
            text119.Text = "Processed by";

            run119.Append(runProperties119);
            run119.Append(text119);

            paragraph81.Append(paragraphProperties81);
            paragraph81.Append(run119);

            tableCell73.Append(tableCellProperties73);
            tableCell73.Append(paragraph81);

            TableCell tableCell74 = new TableCell();

            TableCellProperties tableCellProperties74 = new TableCellProperties();
            TableCellWidth tableCellWidth74 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders74 = new TableCellBorders();
            TopBorder topBorder56 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder76 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder58 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder76 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders74.Append(topBorder56);
            tableCellBorders74.Append(leftBorder76);
            tableCellBorders74.Append(bottomBorder58);
            tableCellBorders74.Append(rightBorder76);
            Shading shading74 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment74 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties74.Append(tableCellWidth74);
            tableCellProperties74.Append(tableCellBorders74);
            tableCellProperties74.Append(shading74);
            tableCellProperties74.Append(tableCellVerticalAlignment74);

            Paragraph paragraph82 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties82 = new ParagraphProperties();
            Justification justification36 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties82 = new ParagraphMarkRunProperties();
            RunFonts runFonts200 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize175 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript175 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties82.Append(runFonts200);
            paragraphMarkRunProperties82.Append(fontSize175);
            paragraphMarkRunProperties82.Append(fontSizeComplexScript175);

            paragraphProperties82.Append(justification36);
            paragraphProperties82.Append(paragraphMarkRunProperties82);

            paragraph82.Append(paragraphProperties82);

            tableCell74.Append(tableCellProperties74);
            tableCell74.Append(paragraph82);

            TableCell tableCell75 = new TableCell();

            TableCellProperties tableCellProperties75 = new TableCellProperties();
            TableCellWidth tableCellWidth75 = new TableCellWidth() { Width = "1008", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders75 = new TableCellBorders();
            TopBorder topBorder57 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder77 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder59 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder77 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders75.Append(topBorder57);
            tableCellBorders75.Append(leftBorder77);
            tableCellBorders75.Append(bottomBorder59);
            tableCellBorders75.Append(rightBorder77);
            Shading shading75 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment75 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties75.Append(tableCellWidth75);
            tableCellProperties75.Append(tableCellBorders75);
            tableCellProperties75.Append(shading75);
            tableCellProperties75.Append(tableCellVerticalAlignment75);

            Paragraph paragraph83 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00D41B28" };

            ParagraphProperties paragraphProperties83 = new ParagraphProperties();
            Indentation indentation1 = new Indentation() { End = "181" };
            Justification justification37 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties83 = new ParagraphMarkRunProperties();
            RunFonts runFonts201 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize176 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript176 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties83.Append(runFonts201);
            paragraphMarkRunProperties83.Append(fontSize176);
            paragraphMarkRunProperties83.Append(fontSizeComplexScript176);

            paragraphProperties83.Append(indentation1);
            paragraphProperties83.Append(justification37);
            paragraphProperties83.Append(paragraphMarkRunProperties83);

            Run run120 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties120 = new RunProperties();
            RunFonts runFonts202 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize177 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript177 = new FontSizeComplexScript() { Val = "16" };

            runProperties120.Append(runFonts202);
            runProperties120.Append(fontSize177);
            runProperties120.Append(fontSizeComplexScript177);
            Text text120 = new Text();
            text120.Text = "Station";

            run120.Append(runProperties120);
            run120.Append(text120);

            paragraph83.Append(paragraphProperties83);
            paragraph83.Append(run120);

            tableCell75.Append(tableCellProperties75);
            tableCell75.Append(paragraph83);

            TableCell tableCell76 = new TableCell();

            TableCellProperties tableCellProperties76 = new TableCellProperties();
            TableCellWidth tableCellWidth76 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders76 = new TableCellBorders();
            TopBorder topBorder58 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder78 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder60 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder78 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders76.Append(topBorder58);
            tableCellBorders76.Append(leftBorder78);
            tableCellBorders76.Append(bottomBorder60);
            tableCellBorders76.Append(rightBorder78);
            Shading shading76 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment76 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties76.Append(tableCellWidth76);
            tableCellProperties76.Append(tableCellBorders76);
            tableCellProperties76.Append(shading76);
            tableCellProperties76.Append(tableCellVerticalAlignment76);

            Paragraph paragraph84 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties84 = new ParagraphProperties();
            Justification justification38 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties84 = new ParagraphMarkRunProperties();
            RunFonts runFonts204 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize179 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript179 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties84.Append(runFonts204);
            paragraphMarkRunProperties84.Append(fontSize179);
            paragraphMarkRunProperties84.Append(fontSizeComplexScript179);

            paragraphProperties84.Append(justification38);
            paragraphProperties84.Append(paragraphMarkRunProperties84);

            Run run122 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties122 = new RunProperties();
            RunFonts runFonts205 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize180 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript180 = new FontSizeComplexScript() { Val = "16" };

            runProperties122.Append(runFonts205);
            runProperties122.Append(fontSize180);
            runProperties122.Append(fontSizeComplexScript180);
            Text text122 = new Text();
            text122.Text = "Time";

            run122.Append(runProperties122);
            run122.Append(text122);

            paragraph84.Append(paragraphProperties84);
            paragraph84.Append(run122);

            tableCell76.Append(tableCellProperties76);
            tableCell76.Append(paragraph84);

            TableCell tableCell77 = new TableCell();

            TableCellProperties tableCellProperties77 = new TableCellProperties();
            TableCellWidth tableCellWidth77 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders77 = new TableCellBorders();
            TopBorder topBorder59 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder79 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder61 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder79 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders77.Append(topBorder59);
            tableCellBorders77.Append(leftBorder79);
            tableCellBorders77.Append(bottomBorder61);
            tableCellBorders77.Append(rightBorder79);
            Shading shading77 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment77 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties77.Append(tableCellWidth77);
            tableCellProperties77.Append(tableCellBorders77);
            tableCellProperties77.Append(shading77);
            tableCellProperties77.Append(tableCellVerticalAlignment77);

            Paragraph paragraph85 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties85 = new ParagraphProperties();
            Justification justification39 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties85 = new ParagraphMarkRunProperties();
            RunFonts runFonts206 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize181 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript181 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties85.Append(runFonts206);
            paragraphMarkRunProperties85.Append(fontSize181);
            paragraphMarkRunProperties85.Append(fontSizeComplexScript181);

            paragraphProperties85.Append(justification39);
            paragraphProperties85.Append(paragraphMarkRunProperties85);

            Run run123 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties123 = new RunProperties();
            RunFonts runFonts207 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize182 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript182 = new FontSizeComplexScript() { Val = "16" };

            runProperties123.Append(runFonts207);
            runProperties123.Append(fontSize182);
            runProperties123.Append(fontSizeComplexScript182);
            Text text123 = new Text();
            text123.Text = "MPN/ 100ml";

            run123.Append(runProperties123);
            run123.Append(text123);

            paragraph85.Append(paragraphProperties85);
            paragraph85.Append(run123);

            Paragraph paragraph86 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties86 = new ParagraphProperties();
            Justification justification40 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties86 = new ParagraphMarkRunProperties();
            RunFonts runFonts208 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize183 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript183 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties86.Append(runFonts208);
            paragraphMarkRunProperties86.Append(fontSize183);
            paragraphMarkRunProperties86.Append(fontSizeComplexScript183);

            paragraphProperties86.Append(justification40);
            paragraphProperties86.Append(paragraphMarkRunProperties86);

            Run run124 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties124 = new RunProperties();
            RunFonts runFonts209 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize184 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript184 = new FontSizeComplexScript() { Val = "16" };

            runProperties124.Append(runFonts209);
            runProperties124.Append(fontSize184);
            runProperties124.Append(fontSizeComplexScript184);
            Text text124 = new Text();
            text124.Text = "Fecal Coliform";

            run124.Append(runProperties124);
            run124.Append(text124);

            paragraph86.Append(paragraphProperties86);
            paragraph86.Append(run124);

            tableCell77.Append(tableCellProperties77);
            tableCell77.Append(paragraph85);
            tableCell77.Append(paragraph86);

            TableCell tableCell78 = new TableCell();

            TableCellProperties tableCellProperties78 = new TableCellProperties();
            TableCellWidth tableCellWidth78 = new TableCellWidth() { Width = "1800", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan3 = new GridSpan() { Val = 3 };

            TableCellBorders tableCellBorders78 = new TableCellBorders();
            TopBorder topBorder60 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder80 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder62 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder80 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders78.Append(topBorder60);
            tableCellBorders78.Append(leftBorder80);
            tableCellBorders78.Append(bottomBorder62);
            tableCellBorders78.Append(rightBorder80);
            Shading shading78 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment78 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties78.Append(tableCellWidth78);
            tableCellProperties78.Append(gridSpan3);
            tableCellProperties78.Append(tableCellBorders78);
            tableCellProperties78.Append(shading78);
            tableCellProperties78.Append(tableCellVerticalAlignment78);

            Paragraph paragraph87 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties87 = new ParagraphProperties();
            Justification justification41 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties87 = new ParagraphMarkRunProperties();
            RunFonts runFonts210 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize185 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript185 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties87.Append(runFonts210);
            paragraphMarkRunProperties87.Append(fontSize185);
            paragraphMarkRunProperties87.Append(fontSizeComplexScript185);

            paragraphProperties87.Append(justification41);
            paragraphProperties87.Append(paragraphMarkRunProperties87);

            Run run125 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties125 = new RunProperties();
            RunFonts runFonts211 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize186 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript186 = new FontSizeComplexScript() { Val = "16" };

            runProperties125.Append(runFonts211);
            runProperties125.Append(fontSize186);
            runProperties125.Append(fontSizeComplexScript186);
            Text text125 = new Text();
            text125.Text = "Positive Tubes";

            run125.Append(runProperties125);
            run125.Append(text125);

            paragraph87.Append(paragraphProperties87);
            paragraph87.Append(run125);

            Paragraph paragraph88 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00D41B28" };

            ParagraphProperties paragraphProperties88 = new ParagraphProperties();
            Justification justification42 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties88 = new ParagraphMarkRunProperties();
            RunFonts runFonts212 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize187 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript187 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties88.Append(runFonts212);
            paragraphMarkRunProperties88.Append(fontSize187);
            paragraphMarkRunProperties88.Append(fontSizeComplexScript187);

            paragraphProperties88.Append(justification42);
            paragraphProperties88.Append(paragraphMarkRunProperties88);

            Run run126 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties126 = new RunProperties();
            RunFonts runFonts213 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize188 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript188 = new FontSizeComplexScript() { Val = "16" };

            runProperties126.Append(runFonts213);
            runProperties126.Append(fontSize188);
            runProperties126.Append(fontSizeComplexScript188);
            Text text126 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text126.Text = "10     1.0  ";

            run126.Append(runProperties126);
            run126.Append(text126);

            Run run127 = new Run() { RsidRunProperties = "00ED7FB2", RsidRunAddition = "00B8087D" };

            RunProperties runProperties127 = new RunProperties();
            RunFonts runFonts214 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize189 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript189 = new FontSizeComplexScript() { Val = "16" };

            runProperties127.Append(runFonts214);
            runProperties127.Append(fontSize189);
            runProperties127.Append(fontSizeComplexScript189);
            Text text127 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text127.Text = "   0.1";

            run127.Append(runProperties127);
            run127.Append(text127);

            paragraph88.Append(paragraphProperties88);
            paragraph88.Append(run126);
            paragraph88.Append(run127);

            tableCell78.Append(tableCellProperties78);
            tableCell78.Append(paragraph87);
            tableCell78.Append(paragraph88);

            TableCell tableCell79 = new TableCell();

            TableCellProperties tableCellProperties79 = new TableCellProperties();
            TableCellWidth tableCellWidth79 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders79 = new TableCellBorders();
            TopBorder topBorder61 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder81 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder63 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder81 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders79.Append(topBorder61);
            tableCellBorders79.Append(leftBorder81);
            tableCellBorders79.Append(bottomBorder63);
            tableCellBorders79.Append(rightBorder81);
            Shading shading79 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment79 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties79.Append(tableCellWidth79);
            tableCellProperties79.Append(tableCellBorders79);
            tableCellProperties79.Append(shading79);
            tableCellProperties79.Append(tableCellVerticalAlignment79);

            Paragraph paragraph89 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties89 = new ParagraphProperties();
            Justification justification43 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties89 = new ParagraphMarkRunProperties();
            RunFonts runFonts215 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize190 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript190 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties89.Append(runFonts215);
            paragraphMarkRunProperties89.Append(fontSize190);
            paragraphMarkRunProperties89.Append(fontSizeComplexScript190);

            paragraphProperties89.Append(justification43);
            paragraphProperties89.Append(paragraphMarkRunProperties89);

            Run run128 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties128 = new RunProperties();
            RunFonts runFonts216 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize191 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript191 = new FontSizeComplexScript() { Val = "16" };

            runProperties128.Append(runFonts216);
            runProperties128.Append(fontSize191);
            runProperties128.Append(fontSizeComplexScript191);
            Text text128 = new Text();
            text128.Text = "Salinity";

            run128.Append(runProperties128);
            run128.Append(text128);

            paragraph89.Append(paragraphProperties89);
            paragraph89.Append(run128);

            Paragraph paragraph90 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00B8087D" };

            ParagraphProperties paragraphProperties90 = new ParagraphProperties();
            Justification justification44 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties90 = new ParagraphMarkRunProperties();
            RunFonts runFonts217 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize192 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript192 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties90.Append(runFonts217);
            paragraphMarkRunProperties90.Append(fontSize192);
            paragraphMarkRunProperties90.Append(fontSizeComplexScript192);

            paragraphProperties90.Append(justification44);
            paragraphProperties90.Append(paragraphMarkRunProperties90);

            Run run129 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties129 = new RunProperties();
            RunFonts runFonts218 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize193 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript193 = new FontSizeComplexScript() { Val = "16" };

            runProperties129.Append(runFonts218);
            runProperties129.Append(fontSize193);
            runProperties129.Append(fontSizeComplexScript193);
            Text text129 = new Text();
            text129.Text = "(ppt)";

            run129.Append(runProperties129);
            run129.Append(text129);

            paragraph90.Append(paragraphProperties90);
            paragraph90.Append(run129);

            tableCell79.Append(tableCellProperties79);
            tableCell79.Append(paragraph89);
            tableCell79.Append(paragraph90);

            TableCell tableCell80 = new TableCell();

            TableCellProperties tableCellProperties80 = new TableCellProperties();
            TableCellWidth tableCellWidth80 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan4 = new GridSpan() { Val = 2 };

            TableCellBorders tableCellBorders80 = new TableCellBorders();
            TopBorder topBorder62 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder82 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder64 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder82 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders80.Append(topBorder62);
            tableCellBorders80.Append(leftBorder82);
            tableCellBorders80.Append(bottomBorder64);
            tableCellBorders80.Append(rightBorder82);
            Shading shading80 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties80.Append(tableCellWidth80);
            tableCellProperties80.Append(gridSpan4);
            tableCellProperties80.Append(tableCellBorders80);
            tableCellProperties80.Append(shading80);

            Paragraph paragraph91 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00D41B28", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00D41B28" };

            ParagraphProperties paragraphProperties91 = new ParagraphProperties();
            Justification justification45 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties91 = new ParagraphMarkRunProperties();
            RunFonts runFonts219 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize194 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript194 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties91.Append(runFonts219);
            paragraphMarkRunProperties91.Append(fontSize194);
            paragraphMarkRunProperties91.Append(fontSizeComplexScript194);

            paragraphProperties91.Append(justification45);
            paragraphProperties91.Append(paragraphMarkRunProperties91);

            Run run130 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties130 = new RunProperties();
            RunFonts runFonts220 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize195 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript195 = new FontSizeComplexScript() { Val = "16" };

            runProperties130.Append(runFonts220);
            runProperties130.Append(fontSize195);
            runProperties130.Append(fontSizeComplexScript195);
            Text text130 = new Text();
            text130.Text = "Temp";

            run130.Append(runProperties130);
            run130.Append(text130);

            paragraph91.Append(paragraphProperties91);
            paragraph91.Append(run130);

            Paragraph paragraph92 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00D41B28" };

            ParagraphProperties paragraphProperties92 = new ParagraphProperties();
            Justification justification46 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties92 = new ParagraphMarkRunProperties();
            RunFonts runFonts221 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize196 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript196 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties92.Append(runFonts221);
            paragraphMarkRunProperties92.Append(fontSize196);
            paragraphMarkRunProperties92.Append(fontSizeComplexScript196);

            paragraphProperties92.Append(justification46);
            paragraphProperties92.Append(paragraphMarkRunProperties92);

            Run run131 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties131 = new RunProperties();
            RunFonts runFonts222 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize197 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript197 = new FontSizeComplexScript() { Val = "16" };

            runProperties131.Append(runFonts222);
            runProperties131.Append(fontSize197);
            runProperties131.Append(fontSizeComplexScript197);
            Text text131 = new Text();
            text131.Text = "(";

            run131.Append(runProperties131);
            run131.Append(text131);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run132 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties132 = new RunProperties();
            RunFonts runFonts223 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize198 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript198 = new FontSizeComplexScript() { Val = "16" };
            VerticalTextAlignment verticalTextAlignment5 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties132.Append(runFonts223);
            runProperties132.Append(fontSize198);
            runProperties132.Append(fontSizeComplexScript198);
            runProperties132.Append(verticalTextAlignment5);
            Text text132 = new Text();
            text132.Text = "o";

            run132.Append(runProperties132);
            run132.Append(text132);

            Run run133 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties133 = new RunProperties();
            RunFonts runFonts224 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize199 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript199 = new FontSizeComplexScript() { Val = "16" };

            runProperties133.Append(runFonts224);
            runProperties133.Append(fontSize199);
            runProperties133.Append(fontSizeComplexScript199);
            Text text133 = new Text();
            text133.Text = "C";

            run133.Append(runProperties133);
            run133.Append(text133);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run134 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties134 = new RunProperties();
            RunFonts runFonts225 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize200 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript200 = new FontSizeComplexScript() { Val = "16" };

            runProperties134.Append(runFonts225);
            runProperties134.Append(fontSize200);
            runProperties134.Append(fontSizeComplexScript200);
            Text text134 = new Text();
            text134.Text = ")";

            run134.Append(runProperties134);
            run134.Append(text134);

            paragraph92.Append(paragraphProperties92);
            paragraph92.Append(run131);
            paragraph92.Append(proofError7);
            paragraph92.Append(run132);
            paragraph92.Append(run133);
            paragraph92.Append(proofError8);
            paragraph92.Append(run134);

            tableCell80.Append(tableCellProperties80);
            tableCell80.Append(paragraph91);
            tableCell80.Append(paragraph92);

            TableCell tableCell81 = new TableCell();

            TableCellProperties tableCellProperties81 = new TableCellProperties();
            TableCellWidth tableCellWidth81 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders81 = new TableCellBorders();
            TopBorder topBorder63 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder83 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder65 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder83 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders81.Append(topBorder63);
            tableCellBorders81.Append(leftBorder83);
            tableCellBorders81.Append(bottomBorder65);
            tableCellBorders81.Append(rightBorder83);
            Shading shading81 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment80 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties81.Append(tableCellWidth81);
            tableCellProperties81.Append(tableCellBorders81);
            tableCellProperties81.Append(shading81);
            tableCellProperties81.Append(tableCellVerticalAlignment80);

            Paragraph paragraph93 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00B8087D", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00286EB4" };

            ParagraphProperties paragraphProperties93 = new ParagraphProperties();
            Justification justification47 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties93 = new ParagraphMarkRunProperties();
            RunFonts runFonts226 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize201 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript201 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties93.Append(runFonts226);
            paragraphMarkRunProperties93.Append(fontSize201);
            paragraphMarkRunProperties93.Append(fontSizeComplexScript201);

            paragraphProperties93.Append(justification47);
            paragraphProperties93.Append(paragraphMarkRunProperties93);

            Run run135 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties135 = new RunProperties();
            RunFonts runFonts227 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize202 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript202 = new FontSizeComplexScript() { Val = "16" };

            runProperties135.Append(runFonts227);
            runProperties135.Append(fontSize202);
            runProperties135.Append(fontSizeComplexScript202);
            Text text135 = new Text();
            text135.Text = "Processed by";

            run135.Append(runProperties135);
            run135.Append(text135);

            paragraph93.Append(paragraphProperties93);
            paragraph93.Append(run135);

            tableCell81.Append(tableCellProperties81);
            tableCell81.Append(paragraph93);

            tableRow6.Append(tableCell67);
            tableRow6.Append(tableCell68);
            tableRow6.Append(tableCell69);
            tableRow6.Append(tableCell70);
            tableRow6.Append(tableCell71);
            tableRow6.Append(tableCell72);
            tableRow6.Append(tableCell73);
            tableRow6.Append(tableCell74);
            tableRow6.Append(tableCell75);
            tableRow6.Append(tableCell76);
            tableRow6.Append(tableCell77);
            tableRow6.Append(tableCell78);
            tableRow6.Append(tableCell79);
            tableRow6.Append(tableCell80);
            tableRow6.Append(tableCell81);

            // TableRowMeasurementRow0

            TableRow tableRow7 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow7, 0, SheetNumber);

            TableRow tableRow8 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow8, 1, SheetNumber);

            TableRow tableRow9 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow9, 2, SheetNumber);

            TableRow tableRow10 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow10, 3, SheetNumber);

            TableRow tableRow11 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow11, 4, SheetNumber);

            TableRow tableRow12 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow12, 5, SheetNumber);

            TableRow tableRow13 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow13, 6, SheetNumber);

            TableRow tableRow14 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow14, 7, SheetNumber);

            TableRow tableRow15 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow15, 8, SheetNumber);

            TableRow tableRow16 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow16, 9, SheetNumber);

            TableRow tableRow17 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow17, 10, SheetNumber);

            TableRow tableRow18 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow18, 11, SheetNumber);

            TableRow tableRow19 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow19, 12, SheetNumber);

            TableRow tableRow20 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow20, 13, SheetNumber);

            TableRow tableRow21 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow21, 14, SheetNumber);

            TableRow tableRow22 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow22, 15, SheetNumber);

            TableRow tableRow23 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow23, 16, SheetNumber);

            TableRow tableRow24 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow24, 17, SheetNumber);

            TableRow tableRow25 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow25, 18, SheetNumber);

            TableRow tableRow26 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00FB2E61", RsidTableRowProperties = "00ED7FB2" };

            FillTableRow(tableRow26, 19, SheetNumber);


            table3.Append(tableProperties3);
            table3.Append(tableGrid3);
            table3.Append(tableRow6);
            table3.Append(tableRow7);
            table3.Append(tableRow8);
            table3.Append(tableRow9);
            table3.Append(tableRow10);
            table3.Append(tableRow11);
            table3.Append(tableRow12);
            table3.Append(tableRow13);
            table3.Append(tableRow14);
            table3.Append(tableRow15);
            table3.Append(tableRow16);
            table3.Append(tableRow17);
            table3.Append(tableRow18);
            table3.Append(tableRow19);
            table3.Append(tableRow20);
            table3.Append(tableRow21);
            table3.Append(tableRow22);
            table3.Append(tableRow23);
            table3.Append(tableRow24);
            table3.Append(tableRow25);
            table3.Append(tableRow26);

            Paragraph paragraph482 = new Paragraph() { RsidParagraphMarkRevision = "009A6852", RsidParagraphAddition = "00B67DA8", RsidRunAdditionDefault = "00B67DA8" };

            ParagraphProperties paragraphProperties482 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties482 = new ParagraphMarkRunProperties();
            RunFonts runFonts728 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties482.Append(runFonts728);

            paragraphProperties482.Append(paragraphMarkRunProperties482);

            Run run248 = new Run();

            RunProperties runProperties248 = new RunProperties();
            RunFonts runFonts729 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Highlight highlight23 = new Highlight() { Val = HighlightColorValues.Yellow };

            runProperties248.Append(runFonts729);
            runProperties248.Append(highlight23);
            Text text248 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text248.Text = " ";

            run248.Append(runProperties248);
            run248.Append(text248);

            paragraph482.Append(paragraphProperties482);
            paragraph482.Append(run248);

            Paragraph paragraph483 = new Paragraph() { RsidParagraphMarkRevision = "002164F2", RsidParagraphAddition = "00004340", RsidRunAdditionDefault = "00D53280" };

            ParagraphProperties paragraphProperties483 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties483 = new ParagraphMarkRunProperties();
            RunFonts runFonts730 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize367 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript367 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties483.Append(runFonts730);
            paragraphMarkRunProperties483.Append(fontSize367);
            paragraphMarkRunProperties483.Append(fontSizeComplexScript367);

            paragraphProperties483.Append(paragraphMarkRunProperties483);

            Run run249 = new Run() { RsidRunProperties = "00D53280" };

            RunProperties runProperties249 = new RunProperties();
            RunFonts runFonts731 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize368 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript368 = new FontSizeComplexScript() { Val = "22" };
            Highlight highlight24 = new Highlight() { Val = HighlightColorValues.Yellow };

            runProperties249.Append(runFonts731);
            runProperties249.Append(fontSize368);
            runProperties249.Append(fontSizeComplexScript368);
            runProperties249.Append(highlight24);
            Text text249 = new Text();
            text249.Text = "Note: All required information as per section 5.10.2 of ISO/IEC 17025 is available from the Laboratory Supervisor.";

            run249.Append(runProperties249);
            run249.Append(text249);

            paragraph483.Append(paragraphProperties483);
            paragraph483.Append(run249);

            Table table4 = new Table();

            TableProperties tableProperties4 = new TableProperties();
            TableWidth tableWidth4 = new TableWidth() { Width = "14760", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation3 = new TableIndentation() { Width = -72, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders3 = new TableBorders();
            TopBorder topBorder300 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder446 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder323 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder446 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder3 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder3 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders3.Append(topBorder300);
            tableBorders3.Append(leftBorder446);
            tableBorders3.Append(bottomBorder323);
            tableBorders3.Append(rightBorder446);
            tableBorders3.Append(insideHorizontalBorder3);
            tableBorders3.Append(insideVerticalBorder3);
            TableLayout tableLayout3 = new TableLayout() { Type = TableLayoutValues.Fixed };
            TableLook tableLook4 = new TableLook() { Val = "01E0" };

            tableProperties4.Append(tableWidth4);
            tableProperties4.Append(tableIndentation3);
            tableProperties4.Append(tableBorders3);
            tableProperties4.Append(tableLayout3);
            tableProperties4.Append(tableLook4);

            TableGrid tableGrid4 = new TableGrid();
            GridColumn gridColumn44 = new GridColumn() { Width = "8224" };
            GridColumn gridColumn45 = new GridColumn() { Width = "236" };
            GridColumn gridColumn46 = new GridColumn() { Width = "2520" };
            GridColumn gridColumn47 = new GridColumn() { Width = "3780" };

            tableGrid4.Append(gridColumn44);
            tableGrid4.Append(gridColumn45);
            tableGrid4.Append(gridColumn46);
            tableGrid4.Append(gridColumn47);

            TableRow tableRow27 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00A34788", RsidTableRowProperties = "00ED7FB2" };

            TableRowProperties tableRowProperties24 = new TableRowProperties();
            TableRowHeight tableRowHeight24 = new TableRowHeight() { Val = (UInt32Value)174U };

            tableRowProperties24.Append(tableRowHeight24);

            TableCell tableCell444 = new TableCell();

            TableCellProperties tableCellProperties444 = new TableCellProperties();
            TableCellWidth tableCellWidth444 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders444 = new TableCellBorders();
            TopBorder topBorder301 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder447 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder324 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder447 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders444.Append(topBorder301);
            tableCellBorders444.Append(leftBorder447);
            tableCellBorders444.Append(bottomBorder324);
            tableCellBorders444.Append(rightBorder447);
            Shading shading444 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment81 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties444.Append(tableCellWidth444);
            tableCellProperties444.Append(tableCellBorders444);
            tableCellProperties444.Append(shading444);
            tableCellProperties444.Append(tableCellVerticalAlignment81);

            Paragraph paragraph484 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00A34788", RsidParagraphProperties = "0012629E", RsidRunAdditionDefault = "00A34788" };

            ParagraphProperties paragraphProperties484 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties484 = new ParagraphMarkRunProperties();
            RunFonts runFonts732 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold13 = new Bold();
            FontSize fontSize369 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript369 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties484.Append(runFonts732);
            paragraphMarkRunProperties484.Append(bold13);
            paragraphMarkRunProperties484.Append(fontSize369);
            paragraphMarkRunProperties484.Append(fontSizeComplexScript369);

            paragraphProperties484.Append(paragraphMarkRunProperties484);

            Run run250 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties250 = new RunProperties();
            RunFonts runFonts733 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold14 = new Bold();
            FontSize fontSize370 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript370 = new FontSizeComplexScript() { Val = "22" };

            runProperties250.Append(runFonts733);
            runProperties250.Append(bold14);
            runProperties250.Append(fontSize370);
            runProperties250.Append(fontSizeComplexScript370);
            Text text250 = new Text();
            text250.Text = "Weather:";

            run250.Append(runProperties250);
            run250.Append(text250);

            Run run251 = new Run() { RsidRunAddition = "0012629E" };

            RunProperties runProperties251 = new RunProperties();
            RunFonts runFonts734 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold15 = new Bold();
            FontSize fontSize371 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript371 = new FontSizeComplexScript() { Val = "22" };

            runProperties251.Append(runFonts734);
            runProperties251.Append(bold15);
            runProperties251.Append(fontSize371);
            runProperties251.Append(fontSizeComplexScript371);
            Text text251 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text251.Text = "  ";

            run251.Append(runProperties251);
            run251.Append(text251);

            Run run252 = new Run() { RsidRunProperties = "0072154A", RsidRunAddition = "0012629E" };

            RunProperties runProperties252 = new RunProperties();
            RunFonts runFonts735 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize372 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript372 = new FontSizeComplexScript() { Val = "20" };

            runProperties252.Append(runFonts735);
            runProperties252.Append(fontSize372);
            runProperties252.Append(fontSizeComplexScript372);
            Text text252 = new Text();
            text252.Text = (labSheetA1Sheet.Weather.Length > 60 ? labSheetA1Sheet.Weather.Substring(0, 55) + " ..." : labSheetA1Sheet.Weather);

            run252.Append(runProperties252);
            run252.Append(text252);

            paragraph484.Append(paragraphProperties484);
            paragraph484.Append(run250);
            paragraph484.Append(run251);
            paragraph484.Append(run252);

            tableCell444.Append(tableCellProperties444);
            tableCell444.Append(paragraph484);

            TableCell tableCell445 = new TableCell();

            TableCellProperties tableCellProperties445 = new TableCellProperties();
            TableCellWidth tableCellWidth445 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders445 = new TableCellBorders();
            TopBorder topBorder302 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder448 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder325 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder448 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders445.Append(topBorder302);
            tableCellBorders445.Append(leftBorder448);
            tableCellBorders445.Append(bottomBorder325);
            tableCellBorders445.Append(rightBorder448);
            Shading shading445 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment82 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties445.Append(tableCellWidth445);
            tableCellProperties445.Append(tableCellBorders445);
            tableCellProperties445.Append(shading445);
            tableCellProperties445.Append(tableCellVerticalAlignment82);

            Paragraph paragraph485 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00A34788", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00A34788" };

            ParagraphProperties paragraphProperties485 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties485 = new ParagraphMarkRunProperties();
            RunFonts runFonts738 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties485.Append(runFonts738);

            paragraphProperties485.Append(paragraphMarkRunProperties485);

            paragraph485.Append(paragraphProperties485);

            tableCell445.Append(tableCellProperties445);
            tableCell445.Append(paragraph485);

            TableCell tableCell446 = new TableCell();

            TableCellProperties tableCellProperties446 = new TableCellProperties();
            TableCellWidth tableCellWidth446 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders446 = new TableCellBorders();
            TopBorder topBorder303 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder449 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder326 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder449 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders446.Append(topBorder303);
            tableCellBorders446.Append(leftBorder449);
            tableCellBorders446.Append(bottomBorder326);
            tableCellBorders446.Append(rightBorder449);
            Shading shading446 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment83 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties446.Append(tableCellWidth446);
            tableCellProperties446.Append(tableCellBorders446);
            tableCellProperties446.Append(shading446);
            tableCellProperties446.Append(tableCellVerticalAlignment83);

            Paragraph paragraph486 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00A34788", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties486 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties486 = new ParagraphMarkRunProperties();
            RunFonts runFonts739 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold16 = new Bold();
            FontSize fontSize375 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript375 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties486.Append(runFonts739);
            paragraphMarkRunProperties486.Append(bold16);
            paragraphMarkRunProperties486.Append(fontSize375);
            paragraphMarkRunProperties486.Append(fontSizeComplexScript375);

            paragraphProperties486.Append(paragraphMarkRunProperties486);

            Run run255 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties255 = new RunProperties();
            RunFonts runFonts740 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold17 = new Bold();
            FontSize fontSize376 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript376 = new FontSizeComplexScript() { Val = "22" };

            runProperties255.Append(runFonts740);
            runProperties255.Append(bold17);
            runProperties255.Append(fontSize376);
            runProperties255.Append(fontSizeComplexScript376);
            Text text255 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text255.Text = "Sample Bottle ";

            run255.Append(runProperties255);
            run255.Append(text255);

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<w:smartTag w:uri=\"urn:schemas-microsoft-com:office:smarttags\" w:element=\"place\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r w:rsidRPr=\"00ED7FB2\"><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\" /><w:b /><w:sz w:val=\"22\" /><w:szCs w:val=\"22\" /></w:rPr><w:t>Lot</w:t></w:r></w:smartTag>");

            Run run256 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties256 = new RunProperties();
            RunFonts runFonts741 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold18 = new Bold();
            FontSize fontSize377 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript377 = new FontSizeComplexScript() { Val = "22" };

            runProperties256.Append(runFonts741);
            runProperties256.Append(bold18);
            runProperties256.Append(fontSize377);
            runProperties256.Append(fontSizeComplexScript377);
            Text text256 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text256.Text = " #";

            run256.Append(runProperties256);
            run256.Append(text256);

            paragraph486.Append(paragraphProperties486);
            paragraph486.Append(run255);
            paragraph486.Append(openXmlUnknownElement3);
            paragraph486.Append(run256);

            tableCell446.Append(tableCellProperties446);
            tableCell446.Append(paragraph486);

            TableCell tableCell447 = new TableCell();

            TableCellProperties tableCellProperties447 = new TableCellProperties();
            TableCellWidth tableCellWidth447 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders447 = new TableCellBorders();
            TopBorder topBorder304 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder450 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder327 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder450 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders447.Append(topBorder304);
            tableCellBorders447.Append(leftBorder450);
            tableCellBorders447.Append(bottomBorder327);
            tableCellBorders447.Append(rightBorder450);
            Shading shading447 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment84 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties447.Append(tableCellWidth447);
            tableCellProperties447.Append(tableCellBorders447);
            tableCellProperties447.Append(shading447);
            tableCellProperties447.Append(tableCellVerticalAlignment84);

            Paragraph paragraph487 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00A34788", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "0012629E" };

            ParagraphProperties paragraphProperties487 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties487 = new ParagraphMarkRunProperties();
            RunFonts runFonts742 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color35 = new Color() { Val = "C0C0C0" };
            FontSize fontSize378 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript378 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties487.Append(runFonts742);
            paragraphMarkRunProperties487.Append(color35);
            paragraphMarkRunProperties487.Append(fontSize378);
            paragraphMarkRunProperties487.Append(fontSizeComplexScript378);

            paragraphProperties487.Append(paragraphMarkRunProperties487);

            Run run257 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties257 = new RunProperties();
            RunFonts runFonts743 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize379 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript379 = new FontSizeComplexScript() { Val = "20" };

            runProperties257.Append(runFonts743);
            runProperties257.Append(fontSize379);
            runProperties257.Append(fontSizeComplexScript379);
            Text text257 = new Text();
            text257.Text = labSheetA1Sheet.SampleBottleLotNumber;

            run257.Append(runProperties257);
            run257.Append(text257);

            paragraph487.Append(paragraphProperties487);
            paragraph487.Append(run257);

            tableCell447.Append(tableCellProperties447);
            tableCell447.Append(paragraph487);

            tableRow27.Append(tableRowProperties24);
            tableRow27.Append(tableCell444);
            tableRow27.Append(tableCell445);
            tableRow27.Append(tableCell446);
            tableRow27.Append(tableCell447);

            TableRow tableRow28 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009A6852", RsidTableRowProperties = "00ED7FB2" };

            TableRowProperties tableRowProperties25 = new TableRowProperties();
            TableRowHeight tableRowHeight25 = new TableRowHeight() { Val = (UInt32Value)271U };

            tableRowProperties25.Append(tableRowHeight25);

            TableCell tableCell448 = new TableCell();

            TableCellProperties tableCellProperties448 = new TableCellProperties();
            TableCellWidth tableCellWidth448 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders448 = new TableCellBorders();
            TopBorder topBorder305 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder451 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder328 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder451 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders448.Append(topBorder305);
            tableCellBorders448.Append(leftBorder451);
            tableCellBorders448.Append(bottomBorder328);
            tableCellBorders448.Append(rightBorder451);
            Shading shading448 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment85 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties448.Append(tableCellWidth448);
            tableCellProperties448.Append(tableCellBorders448);
            tableCellProperties448.Append(shading448);
            tableCellProperties448.Append(tableCellVerticalAlignment85);

            Paragraph paragraph488 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "0012629E", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties488 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties488 = new ParagraphMarkRunProperties();
            RunFonts runFonts746 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold19 = new Bold();
            FontSize fontSize382 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript382 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties488.Append(runFonts746);
            paragraphMarkRunProperties488.Append(bold19);
            paragraphMarkRunProperties488.Append(fontSize382);
            paragraphMarkRunProperties488.Append(fontSizeComplexScript382);

            paragraphProperties488.Append(paragraphMarkRunProperties488);

            Run run260 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties260 = new RunProperties();
            RunFonts runFonts747 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold20 = new Bold();
            FontSize fontSize383 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript383 = new FontSizeComplexScript() { Val = "22" };

            runProperties260.Append(runFonts747);
            runProperties260.Append(bold20);
            runProperties260.Append(fontSize383);
            runProperties260.Append(fontSizeComplexScript383);
            Text text260 = new Text();
            text260.Text = "Comments/ Non Conformances:";

            run260.Append(runProperties260);
            run260.Append(text260);

            Run run261 = new Run() { RsidRunAddition = "0012629E" };

            RunProperties runProperties261 = new RunProperties();
            RunFonts runFonts748 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold21 = new Bold();
            FontSize fontSize384 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript384 = new FontSizeComplexScript() { Val = "22" };

            runProperties261.Append(runFonts748);
            runProperties261.Append(bold21);
            runProperties261.Append(fontSize384);
            runProperties261.Append(fontSizeComplexScript384);
            Text text261 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text261.Text = "  ";

            run261.Append(runProperties261);
            run261.Append(text261);

            Run run262 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties262 = new RunProperties();
            RunFonts runFonts749 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold22 = new Bold();
            FontSize fontSize385 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript385 = new FontSizeComplexScript() { Val = "22" };

            runProperties262.Append(runFonts749);
            runProperties262.Append(bold22);
            runProperties262.Append(fontSize385);
            runProperties262.Append(fontSizeComplexScript385);
            Text text262 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text262.Text = " ";

            run262.Append(runProperties262);
            run262.Append(text262);

            Run run263 = new Run() { RsidRunProperties = "0072154A", RsidRunAddition = "0012629E" };

            RunProperties runProperties263 = new RunProperties();
            RunFonts runFonts750 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize386 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript386 = new FontSizeComplexScript() { Val = "20" };

            runProperties263.Append(runFonts750);
            runProperties263.Append(fontSize386);
            runProperties263.Append(fontSizeComplexScript386);
            Text text263 = new Text();
            text263.Text = (labSheetA1Sheet.RunComment.Length > 40 ? labSheetA1Sheet.RunComment.Substring(0, 35) + " ..." : labSheetA1Sheet.RunComment);

            run263.Append(runProperties263);
            run263.Append(text263);

            paragraph488.Append(paragraphProperties488);
            paragraph488.Append(run260);
            paragraph488.Append(run261);
            paragraph488.Append(run262);
            paragraph488.Append(run263);

            tableCell448.Append(tableCellProperties448);
            tableCell448.Append(paragraph488);

            TableCell tableCell449 = new TableCell();

            TableCellProperties tableCellProperties449 = new TableCellProperties();
            TableCellWidth tableCellWidth449 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders449 = new TableCellBorders();
            TopBorder topBorder306 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder452 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder329 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder452 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders449.Append(topBorder306);
            tableCellBorders449.Append(leftBorder452);
            tableCellBorders449.Append(bottomBorder329);
            tableCellBorders449.Append(rightBorder452);
            Shading shading449 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment86 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties449.Append(tableCellWidth449);
            tableCellProperties449.Append(tableCellBorders449);
            tableCellProperties449.Append(shading449);
            tableCellProperties449.Append(tableCellVerticalAlignment86);

            Paragraph paragraph489 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties489 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties489 = new ParagraphMarkRunProperties();
            RunFonts runFonts753 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Highlight highlight25 = new Highlight() { Val = HighlightColorValues.Yellow };

            paragraphMarkRunProperties489.Append(runFonts753);
            paragraphMarkRunProperties489.Append(highlight25);

            paragraphProperties489.Append(paragraphMarkRunProperties489);

            paragraph489.Append(paragraphProperties489);

            tableCell449.Append(tableCellProperties449);
            tableCell449.Append(paragraph489);

            TableCell tableCell450 = new TableCell();

            TableCellProperties tableCellProperties450 = new TableCellProperties();
            TableCellWidth tableCellWidth450 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders450 = new TableCellBorders();
            TopBorder topBorder307 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder453 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder330 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder453 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders450.Append(topBorder307);
            tableCellBorders450.Append(leftBorder453);
            tableCellBorders450.Append(bottomBorder330);
            tableCellBorders450.Append(rightBorder453);
            Shading shading450 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment87 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties450.Append(tableCellWidth450);
            tableCellProperties450.Append(tableCellBorders450);
            tableCellProperties450.Append(shading450);
            tableCellProperties450.Append(tableCellVerticalAlignment87);

            Paragraph paragraph490 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties490 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties490 = new ParagraphMarkRunProperties();
            RunFonts runFonts754 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold23 = new Bold();
            FontSize fontSize389 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript389 = new FontSizeComplexScript() { Val = "22" };
            Highlight highlight26 = new Highlight() { Val = HighlightColorValues.Yellow };

            paragraphMarkRunProperties490.Append(runFonts754);
            paragraphMarkRunProperties490.Append(bold23);
            paragraphMarkRunProperties490.Append(fontSize389);
            paragraphMarkRunProperties490.Append(fontSizeComplexScript389);
            paragraphMarkRunProperties490.Append(highlight26);

            paragraphProperties490.Append(paragraphMarkRunProperties490);

            Run run266 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties266 = new RunProperties();
            RunFonts runFonts755 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold24 = new Bold();
            FontSize fontSize390 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript390 = new FontSizeComplexScript() { Val = "22" };

            runProperties266.Append(runFonts755);
            runProperties266.Append(bold24);
            runProperties266.Append(fontSize390);
            runProperties266.Append(fontSizeComplexScript390);
            Text text266 = new Text();
            text266.Text = "Salinities Read By";

            run266.Append(runProperties266);
            run266.Append(text266);

            paragraph490.Append(paragraphProperties490);
            paragraph490.Append(run266);

            tableCell450.Append(tableCellProperties450);
            tableCell450.Append(paragraph490);

            TableCell tableCell451 = new TableCell();

            TableCellProperties tableCellProperties451 = new TableCellProperties();
            TableCellWidth tableCellWidth451 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders451 = new TableCellBorders();
            TopBorder topBorder308 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder454 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder331 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder454 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders451.Append(topBorder308);
            tableCellBorders451.Append(leftBorder454);
            tableCellBorders451.Append(bottomBorder331);
            tableCellBorders451.Append(rightBorder454);
            Shading shading451 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment88 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties451.Append(tableCellWidth451);
            tableCellProperties451.Append(tableCellBorders451);
            tableCellProperties451.Append(shading451);
            tableCellProperties451.Append(tableCellVerticalAlignment88);

            Paragraph paragraph491 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "0012629E" };

            ParagraphProperties paragraphProperties491 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties491 = new ParagraphMarkRunProperties();
            RunFonts runFonts756 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color36 = new Color() { Val = "C0C0C0" };
            FontSize fontSize391 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript391 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties491.Append(runFonts756);
            paragraphMarkRunProperties491.Append(color36);
            paragraphMarkRunProperties491.Append(fontSize391);
            paragraphMarkRunProperties491.Append(fontSizeComplexScript391);

            paragraphProperties491.Append(paragraphMarkRunProperties491);

            Run run267 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties267 = new RunProperties();
            RunFonts runFonts757 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize392 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript392 = new FontSizeComplexScript() { Val = "20" };

            runProperties267.Append(runFonts757);
            runProperties267.Append(fontSize392);
            runProperties267.Append(fontSizeComplexScript392);
            Text text267 = new Text();
            DateTime dateTimeSalinities = new DateTime(int.Parse(labSheetA1Sheet.SalinitiesReadYear), int.Parse(labSheetA1Sheet.SalinitiesReadMonth), int.Parse(labSheetA1Sheet.SalinitiesReadDay));
            text267.Text = labSheetA1Sheet.SalinitiesReadBy + "        " + dateTimeSalinities.ToString("yyyy MMMM dd");

            run267.Append(runProperties267);
            run267.Append(text267);

            paragraph491.Append(paragraphProperties491);
            paragraph491.Append(run267);

            tableCell451.Append(tableCellProperties451);
            tableCell451.Append(paragraph491);

            tableRow28.Append(tableRowProperties25);
            tableRow28.Append(tableCell448);
            tableRow28.Append(tableCell449);
            tableRow28.Append(tableCell450);
            tableRow28.Append(tableCell451);

            TableRow tableRow29 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009A6852", RsidTableRowProperties = "00ED7FB2" };

            TableRowProperties tableRowProperties26 = new TableRowProperties();
            TableRowHeight tableRowHeight26 = new TableRowHeight() { Val = (UInt32Value)271U };

            tableRowProperties26.Append(tableRowHeight26);

            TableCell tableCell452 = new TableCell();

            TableCellProperties tableCellProperties452 = new TableCellProperties();
            TableCellWidth tableCellWidth452 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders452 = new TableCellBorders();
            TopBorder topBorder309 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder455 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder332 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder455 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders452.Append(topBorder309);
            tableCellBorders452.Append(leftBorder455);
            tableCellBorders452.Append(bottomBorder332);
            tableCellBorders452.Append(rightBorder455);
            Shading shading452 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment89 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties452.Append(tableCellWidth452);
            tableCellProperties452.Append(tableCellBorders452);
            tableCellProperties452.Append(shading452);
            tableCellProperties452.Append(tableCellVerticalAlignment89);

            Paragraph paragraph492 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties492 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties492 = new ParagraphMarkRunProperties();
            RunFonts runFonts760 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties492.Append(runFonts760);

            paragraphProperties492.Append(paragraphMarkRunProperties492);

            Run run270 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties270 = new RunProperties();
            RunFonts runFonts761 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties270.Append(runFonts761);
            Text text270 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text270.Text = "If results are 0-0-0, the technician is not required to enter <2. ";

            run270.Append(runProperties270);
            run270.Append(text270);

            paragraph492.Append(paragraphProperties492);
            paragraph492.Append(run270);

            tableCell452.Append(tableCellProperties452);
            tableCell452.Append(paragraph492);

            TableCell tableCell453 = new TableCell();

            TableCellProperties tableCellProperties453 = new TableCellProperties();
            TableCellWidth tableCellWidth453 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders453 = new TableCellBorders();
            TopBorder topBorder310 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder456 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder333 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder456 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders453.Append(topBorder310);
            tableCellBorders453.Append(leftBorder456);
            tableCellBorders453.Append(bottomBorder333);
            tableCellBorders453.Append(rightBorder456);
            Shading shading453 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment90 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties453.Append(tableCellWidth453);
            tableCellProperties453.Append(tableCellBorders453);
            tableCellProperties453.Append(shading453);
            tableCellProperties453.Append(tableCellVerticalAlignment90);

            Paragraph paragraph493 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties493 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties493 = new ParagraphMarkRunProperties();
            RunFonts runFonts762 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties493.Append(runFonts762);

            paragraphProperties493.Append(paragraphMarkRunProperties493);

            paragraph493.Append(paragraphProperties493);

            tableCell453.Append(tableCellProperties453);
            tableCell453.Append(paragraph493);

            TableCell tableCell454 = new TableCell();

            TableCellProperties tableCellProperties454 = new TableCellProperties();
            TableCellWidth tableCellWidth454 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders454 = new TableCellBorders();
            TopBorder topBorder311 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            LeftBorder leftBorder457 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder334 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder457 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders454.Append(topBorder311);
            tableCellBorders454.Append(leftBorder457);
            tableCellBorders454.Append(bottomBorder334);
            tableCellBorders454.Append(rightBorder457);
            Shading shading454 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment91 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties454.Append(tableCellWidth454);
            tableCellProperties454.Append(tableCellBorders454);
            tableCellProperties454.Append(shading454);
            tableCellProperties454.Append(tableCellVerticalAlignment91);

            Paragraph paragraph494 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties494 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties494 = new ParagraphMarkRunProperties();
            RunFonts runFonts763 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold25 = new Bold();
            FontSize fontSize395 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript395 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties494.Append(runFonts763);
            paragraphMarkRunProperties494.Append(bold25);
            paragraphMarkRunProperties494.Append(fontSize395);
            paragraphMarkRunProperties494.Append(fontSizeComplexScript395);

            paragraphProperties494.Append(paragraphMarkRunProperties494);

            Run run271 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties271 = new RunProperties();
            RunFonts runFonts764 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold26 = new Bold();
            FontSize fontSize396 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript396 = new FontSizeComplexScript() { Val = "22" };

            runProperties271.Append(runFonts764);
            runProperties271.Append(bold26);
            runProperties271.Append(fontSize396);
            runProperties271.Append(fontSizeComplexScript396);
            Text text271 = new Text();
            text271.Text = "Results Recorded By";

            run271.Append(runProperties271);
            run271.Append(text271);

            paragraph494.Append(paragraphProperties494);
            paragraph494.Append(run271);

            tableCell454.Append(tableCellProperties454);
            tableCell454.Append(paragraph494);

            TableCell tableCell455 = new TableCell();

            TableCellProperties tableCellProperties455 = new TableCellProperties();
            TableCellWidth tableCellWidth455 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders455 = new TableCellBorders();
            TopBorder topBorder312 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder458 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder335 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder458 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders455.Append(topBorder312);
            tableCellBorders455.Append(leftBorder458);
            tableCellBorders455.Append(bottomBorder335);
            tableCellBorders455.Append(rightBorder458);
            Shading shading455 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment92 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties455.Append(tableCellWidth455);
            tableCellProperties455.Append(tableCellBorders455);
            tableCellProperties455.Append(shading455);
            tableCellProperties455.Append(tableCellVerticalAlignment92);

            Paragraph paragraph495 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "0012629E" };

            ParagraphProperties paragraphProperties495 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties495 = new ParagraphMarkRunProperties();
            RunFonts runFonts765 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color37 = new Color() { Val = "C0C0C0" };
            FontSize fontSize397 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript397 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties495.Append(runFonts765);
            paragraphMarkRunProperties495.Append(color37);
            paragraphMarkRunProperties495.Append(fontSize397);
            paragraphMarkRunProperties495.Append(fontSizeComplexScript397);

            paragraphProperties495.Append(paragraphMarkRunProperties495);

            Run run272 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties272 = new RunProperties();
            RunFonts runFonts766 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize398 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript398 = new FontSizeComplexScript() { Val = "20" };

            runProperties272.Append(runFonts766);
            runProperties272.Append(fontSize398);
            runProperties272.Append(fontSizeComplexScript398);
            Text text272 = new Text();
            DateTime dateTimeResultsRecorded = new DateTime(int.Parse(labSheetA1Sheet.ResultsRecordedYear), int.Parse(labSheetA1Sheet.ResultsRecordedMonth), int.Parse(labSheetA1Sheet.ResultsRecordedDay));
            text272.Text = labSheetA1Sheet.ResultsRecordedBy + "        " + dateTimeResultsRecorded.ToString("yyyy MMMM dd");

            run272.Append(runProperties272);
            run272.Append(text272);

            paragraph495.Append(paragraphProperties495);
            paragraph495.Append(run272);

            tableCell455.Append(tableCellProperties455);
            tableCell455.Append(paragraph495);

            tableRow29.Append(tableRowProperties26);
            tableRow29.Append(tableCell452);
            tableRow29.Append(tableCell453);
            tableRow29.Append(tableCell454);
            tableRow29.Append(tableCell455);

            TableRow tableRow30 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009A6852", RsidTableRowProperties = "00ED7FB2" };

            TableRowProperties tableRowProperties27 = new TableRowProperties();
            TableRowHeight tableRowHeight27 = new TableRowHeight() { Val = (UInt32Value)254U };

            tableRowProperties27.Append(tableRowHeight27);

            TableCell tableCell456 = new TableCell();

            TableCellProperties tableCellProperties456 = new TableCellProperties();
            TableCellWidth tableCellWidth456 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders456 = new TableCellBorders();
            TopBorder topBorder313 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder459 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder336 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder459 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders456.Append(topBorder313);
            tableCellBorders456.Append(leftBorder459);
            tableCellBorders456.Append(bottomBorder336);
            tableCellBorders456.Append(rightBorder459);
            Shading shading456 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment93 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties456.Append(tableCellWidth456);
            tableCellProperties456.Append(tableCellBorders456);
            tableCellProperties456.Append(shading456);
            tableCellProperties456.Append(tableCellVerticalAlignment93);

            Paragraph paragraph496 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties496 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties496 = new ParagraphMarkRunProperties();
            RunFonts runFonts769 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties496.Append(runFonts769);

            paragraphProperties496.Append(paragraphMarkRunProperties496);

            Run run275 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties275 = new RunProperties();
            RunFonts runFonts770 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color38 = new Color() { Val = "C0C0C0" };
            FontSize fontSize401 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript401 = new FontSizeComplexScript() { Val = "16" };

            runProperties275.Append(runFonts770);
            runProperties275.Append(color38);
            runProperties275.Append(fontSize401);
            runProperties275.Append(fontSizeComplexScript401);
            Text text275 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text275.Text = "                                   Activity such as boating, weather, debris in water, birds, etc";

            run275.Append(runProperties275);
            run275.Append(text275);

            paragraph496.Append(paragraphProperties496);
            paragraph496.Append(run275);

            tableCell456.Append(tableCellProperties456);
            tableCell456.Append(paragraph496);

            TableCell tableCell457 = new TableCell();

            TableCellProperties tableCellProperties457 = new TableCellProperties();
            TableCellWidth tableCellWidth457 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders457 = new TableCellBorders();
            TopBorder topBorder314 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder460 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder337 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder460 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders457.Append(topBorder314);
            tableCellBorders457.Append(leftBorder460);
            tableCellBorders457.Append(bottomBorder337);
            tableCellBorders457.Append(rightBorder460);
            Shading shading457 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment94 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties457.Append(tableCellWidth457);
            tableCellProperties457.Append(tableCellBorders457);
            tableCellProperties457.Append(shading457);
            tableCellProperties457.Append(tableCellVerticalAlignment94);

            Paragraph paragraph497 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties497 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties497 = new ParagraphMarkRunProperties();
            RunFonts runFonts786 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties497.Append(runFonts786);

            paragraphProperties497.Append(paragraphMarkRunProperties497);

            paragraph497.Append(paragraphProperties497);

            tableCell457.Append(tableCellProperties457);
            tableCell457.Append(paragraph497);

            TableCell tableCell458 = new TableCell();

            TableCellProperties tableCellProperties458 = new TableCellProperties();
            TableCellWidth tableCellWidth458 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders458 = new TableCellBorders();
            TopBorder topBorder315 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder461 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder338 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder461 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders458.Append(topBorder315);
            tableCellBorders458.Append(leftBorder461);
            tableCellBorders458.Append(bottomBorder338);
            tableCellBorders458.Append(rightBorder461);
            Shading shading458 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment95 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties458.Append(tableCellWidth458);
            tableCellProperties458.Append(tableCellBorders458);
            tableCellProperties458.Append(shading458);
            tableCellProperties458.Append(tableCellVerticalAlignment95);

            Paragraph paragraph498 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties498 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties498 = new ParagraphMarkRunProperties();
            RunFonts runFonts787 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold27 = new Bold();
            FontSize fontSize417 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript417 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties498.Append(runFonts787);
            paragraphMarkRunProperties498.Append(bold27);
            paragraphMarkRunProperties498.Append(fontSize417);
            paragraphMarkRunProperties498.Append(fontSizeComplexScript417);

            paragraphProperties498.Append(paragraphMarkRunProperties498);

            Run run291 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties291 = new RunProperties();
            RunFonts runFonts788 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold28 = new Bold();
            FontSize fontSize418 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript418 = new FontSizeComplexScript() { Val = "22" };

            runProperties291.Append(runFonts788);
            runProperties291.Append(bold28);
            runProperties291.Append(fontSize418);
            runProperties291.Append(fontSizeComplexScript418);
            Text text291 = new Text();
            text291.Text = "Approval Date";

            run291.Append(runProperties291);
            run291.Append(text291);

            paragraph498.Append(paragraphProperties498);
            paragraph498.Append(run291);

            tableCell458.Append(tableCellProperties458);
            tableCell458.Append(paragraph498);

            TableCell tableCell459 = new TableCell();

            TableCellProperties tableCellProperties459 = new TableCellProperties();
            TableCellWidth tableCellWidth459 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders459 = new TableCellBorders();
            TopBorder topBorder316 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder462 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder339 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder462 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders459.Append(topBorder316);
            tableCellBorders459.Append(leftBorder462);
            tableCellBorders459.Append(bottomBorder339);
            tableCellBorders459.Append(rightBorder462);
            Shading shading459 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment96 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties459.Append(tableCellWidth459);
            tableCellProperties459.Append(tableCellBorders459);
            tableCellProperties459.Append(shading459);
            tableCellProperties459.Append(tableCellVerticalAlignment96);

            Paragraph paragraph499 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "0012629E", RsidRunAdditionDefault = "0012629E" };

            ParagraphProperties paragraphProperties499 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties499 = new ParagraphMarkRunProperties();
            RunFonts runFonts789 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color54 = new Color() { Val = "C0C0C0" };
            FontSize fontSize419 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript419 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties499.Append(runFonts789);
            paragraphMarkRunProperties499.Append(color54);
            paragraphMarkRunProperties499.Append(fontSize419);
            paragraphMarkRunProperties499.Append(fontSizeComplexScript419);

            paragraphProperties499.Append(paragraphMarkRunProperties499);

            Run run292 = new Run();

            RunProperties runProperties292 = new RunProperties();
            RunFonts runFonts790 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize420 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript420 = new FontSizeComplexScript() { Val = "20" };

            runProperties292.Append(runFonts790);
            runProperties292.Append(fontSize420);
            runProperties292.Append(fontSizeComplexScript420);
            Text text292 = new Text();
            text292.Text = ((DateTime)labSheetModel.ApprovedOrRejectedDateTime).ToString("yyyy MMMM dd");

            run292.Append(runProperties292);
            run292.Append(text292);

            paragraph499.Append(paragraphProperties499);
            paragraph499.Append(run292);

            tableCell459.Append(tableCellProperties459);
            tableCell459.Append(paragraph499);

            tableRow30.Append(tableRowProperties27);
            tableRow30.Append(tableCell456);
            tableRow30.Append(tableCell457);
            tableRow30.Append(tableCell458);
            tableRow30.Append(tableCell459);

            TableRow tableRow31 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009A6852", RsidTableRowProperties = "00D53280" };

            TableRowProperties tableRowProperties28 = new TableRowProperties();
            TableRowHeight tableRowHeight28 = new TableRowHeight() { Val = (UInt32Value)164U };

            tableRowProperties28.Append(tableRowHeight28);

            TableCell tableCell460 = new TableCell();

            TableCellProperties tableCellProperties460 = new TableCellProperties();
            TableCellWidth tableCellWidth460 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders460 = new TableCellBorders();
            TopBorder topBorder317 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder463 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder340 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder463 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders460.Append(topBorder317);
            tableCellBorders460.Append(leftBorder463);
            tableCellBorders460.Append(bottomBorder340);
            tableCellBorders460.Append(rightBorder463);
            Shading shading460 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment97 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties460.Append(tableCellWidth460);
            tableCellProperties460.Append(tableCellBorders460);
            tableCellProperties460.Append(shading460);
            tableCellProperties460.Append(tableCellVerticalAlignment97);

            Paragraph paragraph500 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties500 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties500 = new ParagraphMarkRunProperties();
            RunFonts runFonts791 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize421 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript421 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties500.Append(runFonts791);
            paragraphMarkRunProperties500.Append(fontSize421);
            paragraphMarkRunProperties500.Append(fontSizeComplexScript421);

            paragraphProperties500.Append(paragraphMarkRunProperties500);

            paragraph500.Append(paragraphProperties500);

            tableCell460.Append(tableCellProperties460);
            tableCell460.Append(paragraph500);

            TableCell tableCell461 = new TableCell();

            TableCellProperties tableCellProperties461 = new TableCellProperties();
            TableCellWidth tableCellWidth461 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders461 = new TableCellBorders();
            TopBorder topBorder318 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder464 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder341 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder464 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders461.Append(topBorder318);
            tableCellBorders461.Append(leftBorder464);
            tableCellBorders461.Append(bottomBorder341);
            tableCellBorders461.Append(rightBorder464);
            Shading shading461 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment98 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties461.Append(tableCellWidth461);
            tableCellProperties461.Append(tableCellBorders461);
            tableCellProperties461.Append(shading461);
            tableCellProperties461.Append(tableCellVerticalAlignment98);

            Paragraph paragraph501 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties501 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties501 = new ParagraphMarkRunProperties();
            RunFonts runFonts792 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties501.Append(runFonts792);

            paragraphProperties501.Append(paragraphMarkRunProperties501);

            paragraph501.Append(paragraphProperties501);

            tableCell461.Append(tableCellProperties461);
            tableCell461.Append(paragraph501);

            TableCell tableCell462 = new TableCell();

            TableCellProperties tableCellProperties462 = new TableCellProperties();
            TableCellWidth tableCellWidth462 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders462 = new TableCellBorders();
            TopBorder topBorder319 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder465 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder342 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder465 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders462.Append(topBorder319);
            tableCellBorders462.Append(leftBorder465);
            tableCellBorders462.Append(bottomBorder342);
            tableCellBorders462.Append(rightBorder465);
            Shading shading462 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment99 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties462.Append(tableCellWidth462);
            tableCellProperties462.Append(tableCellBorders462);
            tableCellProperties462.Append(shading462);
            tableCellProperties462.Append(tableCellVerticalAlignment99);

            Paragraph paragraph502 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties502 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties502 = new ParagraphMarkRunProperties();
            RunFonts runFonts793 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold29 = new Bold();
            FontSize fontSize422 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript422 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties502.Append(runFonts793);
            paragraphMarkRunProperties502.Append(bold29);
            paragraphMarkRunProperties502.Append(fontSize422);
            paragraphMarkRunProperties502.Append(fontSizeComplexScript422);

            paragraphProperties502.Append(paragraphMarkRunProperties502);

            Run run293 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties293 = new RunProperties();
            RunFonts runFonts794 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold30 = new Bold();
            FontSize fontSize423 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript423 = new FontSizeComplexScript() { Val = "22" };

            runProperties293.Append(runFonts794);
            runProperties293.Append(bold30);
            runProperties293.Append(fontSize423);
            runProperties293.Append(fontSizeComplexScript423);
            Text text293 = new Text();
            text293.Text = "Supervisor Signature";

            run293.Append(runProperties293);
            run293.Append(text293);

            paragraph502.Append(paragraphProperties502);
            paragraph502.Append(run293);

            tableCell462.Append(tableCellProperties462);
            tableCell462.Append(paragraph502);

            TableCell tableCell463 = new TableCell();

            TableCellProperties tableCellProperties463 = new TableCellProperties();
            TableCellWidth tableCellWidth463 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders463 = new TableCellBorders();
            TopBorder topBorder320 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder466 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder343 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder466 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders463.Append(topBorder320);
            tableCellBorders463.Append(leftBorder466);
            tableCellBorders463.Append(bottomBorder343);
            tableCellBorders463.Append(rightBorder466);
            Shading shading463 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties463.Append(tableCellWidth463);
            tableCellProperties463.Append(tableCellBorders463);
            tableCellProperties463.Append(shading463);

            Paragraph paragraph503 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties503 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties503 = new ParagraphMarkRunProperties();
            RunFonts runFonts795 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            //Color color55 = new Color() { Val = "C0C0C0" };
            FontSize fontSize424 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript424 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties503.Append(runFonts795);
            //paragraphMarkRunProperties503.Append(color55);
            paragraphMarkRunProperties503.Append(fontSize424);
            paragraphMarkRunProperties503.Append(fontSizeComplexScript424);

            paragraphProperties503.Append(paragraphMarkRunProperties503);

            Run run294 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties294 = new RunProperties();
            RunFonts runFonts796 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            //Color color56 = new Color() { Val = "C0C0C0" };
            FontSize fontSize425 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript425 = new FontSizeComplexScript() { Val = "20" };

            runProperties294.Append(runFonts796);
            //runProperties294.Append(color56);
            runProperties294.Append(fontSize425);
            runProperties294.Append(fontSizeComplexScript425);
            Text text294 = new Text();
            text294.Text = labSheetModel.ApprovedOrRejectedByContactTVText;

            run294.Append(runProperties294);
            run294.Append(text294);

            paragraph503.Append(paragraphProperties503);
            paragraph503.Append(run294);

            tableCell463.Append(tableCellProperties463);
            tableCell463.Append(paragraph503);

            tableRow31.Append(tableRowProperties28);
            tableRow31.Append(tableCell460);
            tableRow31.Append(tableCell461);
            tableRow31.Append(tableCell462);
            tableRow31.Append(tableCell463);

            table4.Append(tableProperties4);
            table4.Append(tableGrid4);
            table4.Append(tableRow27);
            table4.Append(tableRow28);
            table4.Append(tableRow29);
            table4.Append(tableRow30);
            table4.Append(tableRow31);
            Paragraph paragraph504 = new Paragraph() { RsidParagraphMarkRevision = "00584876", RsidParagraphAddition = "00584876", RsidParagraphProperties = "002164F2", RsidRunAdditionDefault = "00584876" };

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "00584876", RsidR = "00584876", RsidSect = "00D53280" };
            HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "rId7" };
            FooterReference footerReference1 = new FooterReference() { Type = HeaderFooterValues.Default, Id = "rId8" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)15842U, Height = (UInt32Value)12242U, Orient = PageOrientationValues.Landscape, Code = (UInt16Value)1U };
            PageMargin pageMargin1 = new PageMargin() { Top = 112, Right = (UInt32Value)170U, Bottom = 284, Left = (UInt32Value)567U, Header = (UInt32Value)284U, Footer = (UInt32Value)431U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(headerReference1);
            sectionProperties1.Append(footerReference1);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body.Append(paragraph1);
            body.Append(table1);
            body.Append(paragraph10);
            body.Append(table2);
            body.Append(paragraph70);
            body.Append(table3);
            body.Append(paragraph482);
            body.Append(paragraph483);
            body.Append(table4);
            body.Append(paragraph504);
            body.Append(sectionProperties1);

        }
    }
}