using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
using CSSPWebToolsDB.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class GeneratedClassFCFormDocx
    {
        private void FillTableRow(TableRow tableRow, int index, int SheetNumber)
        {
            if (SheetNumber == 2)
            {
                index = index + 34;
            }
            int A1MeasurementListCount = labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.SampleType != CSSPLabSheetParserDLL.Models.SampleTypeEnum.Duplicate).Count();
            int DuplicateIndex = -1;
            if (labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.SampleType == SampleTypeEnum.Duplicate).Count() > 0)
            {
                DuplicateIndex = labSheetA1Sheet.LabSheetA1MeasurementList.Count - 1;
            }

            BorderValues topBorderValues = BorderValues.Single;
            if (index % 5 == 0)
            {
                topBorderValues = BorderValues.Double;
            }
            TableRowProperties tableRowProperties4 = new TableRowProperties();
            TableRowHeight tableRowHeight4 = new TableRowHeight() { Val = (UInt32Value)270U };

            tableRowProperties4.Append(tableRowHeight4);

            TableCell tableCell82 = new TableCell();

            TableCellProperties tableCellProperties82 = new TableCellProperties();
            TableCellWidth tableCellWidth82 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders82 = new TableCellBorders();
            TopBorder topBorder64 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder84 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder66 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder84 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders82.Append(topBorder64);
            tableCellBorders82.Append(leftBorder84);
            tableCellBorders82.Append(bottomBorder66);
            tableCellBorders82.Append(rightBorder84);
            Shading shading82 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties82.Append(tableCellWidth82);
            tableCellProperties82.Append(tableCellBorders82);
            tableCellProperties82.Append(shading82);

            Paragraph paragraph94 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties94 = new ParagraphProperties();
            Justification justification48 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties94 = new ParagraphMarkRunProperties();
            RunFonts runFonts228 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties94.Append(runFonts228);

            paragraphProperties94.Append(justification48);
            paragraphProperties94.Append(paragraphMarkRunProperties94);

            Run run136 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties136 = new RunProperties();
            RunFonts runFonts229 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize203 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript203 = new FontSizeComplexScript() { Val = "20" };

            runProperties136.Append(runFonts229);
            runProperties136.Append(fontSize203);
            runProperties136.Append(fontSizeComplexScript203);
            Text text136 = new Text();
            text136.Text = (A1MeasurementListCount > index ? labSheetA1Sheet.LabSheetA1MeasurementList[index].Site : "");

            run136.Append(runProperties136);
            run136.Append(text136);

            paragraph94.Append(paragraphProperties94);
            paragraph94.Append(run136);

            tableCell82.Append(tableCellProperties82);
            tableCell82.Append(paragraph94);

            TableCell tableCell83 = new TableCell();

            TableCellProperties tableCellProperties83 = new TableCellProperties();
            TableCellWidth tableCellWidth83 = new TableCellWidth() { Width = "864", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders83 = new TableCellBorders();
            TopBorder topBorder65 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder85 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder67 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder85 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders83.Append(topBorder65);
            tableCellBorders83.Append(leftBorder85);
            tableCellBorders83.Append(bottomBorder67);
            tableCellBorders83.Append(rightBorder85);
            Shading shading83 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties83.Append(tableCellWidth83);
            tableCellProperties83.Append(tableCellBorders83);
            tableCellProperties83.Append(shading83);

            Paragraph paragraph95 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties95 = new ParagraphProperties();
            Justification justification49 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties95 = new ParagraphMarkRunProperties();
            RunFonts runFonts232 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties95.Append(runFonts232);

            paragraphProperties95.Append(justification49);
            paragraphProperties95.Append(paragraphMarkRunProperties95);

            Run run139 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties139 = new RunProperties();
            RunFonts runFonts233 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize206 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript206 = new FontSizeComplexScript() { Val = "20" };

            runProperties139.Append(runFonts233);
            runProperties139.Append(fontSize206);
            runProperties139.Append(fontSizeComplexScript206);
            Text text139 = new Text();
            if (labSheetA1Sheet.LabSheetA1MeasurementList[index].Time == null)
            {
                text139.Text = (A1MeasurementListCount > index ? "" : "");
            }
            else
            {
                text139.Text = (A1MeasurementListCount > index ? ((DateTime)labSheetA1Sheet.LabSheetA1MeasurementList[index].Time).ToString("hh:mm") : "");
            }

            run139.Append(runProperties139);
            run139.Append(text139);

            paragraph95.Append(paragraphProperties95);
            paragraph95.Append(run139);

            tableCell83.Append(tableCellProperties83);
            tableCell83.Append(paragraph95);

            TableCell tableCell84 = new TableCell();

            TableCellProperties tableCellProperties84 = new TableCellProperties();
            TableCellWidth tableCellWidth84 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders84 = new TableCellBorders();
            TopBorder topBorder66 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder86 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder68 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder86 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders84.Append(topBorder66);
            tableCellBorders84.Append(leftBorder86);
            tableCellBorders84.Append(bottomBorder68);
            tableCellBorders84.Append(rightBorder86);
            Shading shading84 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties84.Append(tableCellWidth84);
            tableCellProperties84.Append(tableCellBorders84);
            tableCellProperties84.Append(shading84);

            Paragraph paragraph96 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties96 = new ParagraphProperties();
            Justification justification50 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties96 = new ParagraphMarkRunProperties();
            RunFonts runFonts236 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties96.Append(runFonts236);

            paragraphProperties96.Append(justification50);
            paragraphProperties96.Append(paragraphMarkRunProperties96);

            Run run142 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties142 = new RunProperties();
            RunFonts runFonts237 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize209 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript209 = new FontSizeComplexScript() { Val = "20" };

            runProperties142.Append(runFonts237);
            runProperties142.Append(fontSize209);
            runProperties142.Append(fontSizeComplexScript209);
            Text text142 = new Text();
            if (labSheetA1Sheet.LabSheetA1MeasurementList[index].MPN == null)
            {
                text142.Text = (A1MeasurementListCount > index ? "" : "");
            }
            else
            {
                text142.Text = (A1MeasurementListCount > index ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[index].MPN).ToString() : "");
            }

            run142.Append(runProperties142);
            run142.Append(text142);

            paragraph96.Append(paragraphProperties96);
            paragraph96.Append(run142);

            tableCell84.Append(tableCellProperties84);
            tableCell84.Append(paragraph96);

            TableCell tableCell85 = new TableCell();

            TableCellProperties tableCellProperties85 = new TableCellProperties();
            TableCellWidth tableCellWidth85 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders85 = new TableCellBorders();
            TopBorder topBorder67 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder87 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder87 = new RightBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders85.Append(topBorder67);
            tableCellBorders85.Append(leftBorder87);
            tableCellBorders85.Append(rightBorder87);
            Shading shading85 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties85.Append(tableCellWidth85);
            tableCellProperties85.Append(tableCellBorders85);
            tableCellProperties85.Append(shading85);

            Paragraph paragraph97 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties97 = new ParagraphProperties();
            Justification justification51 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties97 = new ParagraphMarkRunProperties();
            RunFonts runFonts240 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties97.Append(runFonts240);

            paragraphProperties97.Append(justification51);
            paragraphProperties97.Append(paragraphMarkRunProperties97);

            Run run145 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties145 = new RunProperties();
            RunFonts runFonts241 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize212 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript212 = new FontSizeComplexScript() { Val = "20" };

            runProperties145.Append(runFonts241);
            runProperties145.Append(fontSize212);
            runProperties145.Append(fontSizeComplexScript212);
            Text text145 = new Text();
            if (labSheetA1Sheet.LabSheetA1MeasurementList[index].Tube10 == null)
            {
                text145.Text = (A1MeasurementListCount > index ? "" : "");
            }
            else
            {
                text145.Text = (A1MeasurementListCount > index ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[index].Tube10).ToString() : "");
            }

            run145.Append(runProperties145);
            run145.Append(text145);

            paragraph97.Append(paragraphProperties97);
            paragraph97.Append(run145);

            tableCell85.Append(tableCellProperties85);
            tableCell85.Append(paragraph97);

            TableCell tableCell86 = new TableCell();

            TableCellProperties tableCellProperties86 = new TableCellProperties();
            TableCellWidth tableCellWidth86 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders86 = new TableCellBorders();
            TopBorder topBorder68 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder88 = new LeftBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder88 = new RightBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders86.Append(topBorder68);
            tableCellBorders86.Append(leftBorder88);
            tableCellBorders86.Append(rightBorder88);
            Shading shading86 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties86.Append(tableCellWidth86);
            tableCellProperties86.Append(tableCellBorders86);
            tableCellProperties86.Append(shading86);

            Paragraph paragraph98 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties98 = new ParagraphProperties();
            Justification justification52 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties98 = new ParagraphMarkRunProperties();
            RunFonts runFonts244 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties98.Append(runFonts244);

            paragraphProperties98.Append(justification52);
            paragraphProperties98.Append(paragraphMarkRunProperties98);

            Run run148 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties148 = new RunProperties();
            RunFonts runFonts245 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize215 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript215 = new FontSizeComplexScript() { Val = "20" };

            runProperties148.Append(runFonts245);
            runProperties148.Append(fontSize215);
            runProperties148.Append(fontSizeComplexScript215);
            Text text148 = new Text();

            if (labSheetA1Sheet.LabSheetA1MeasurementList[index].Tube1_0 == null)
            {
                text148.Text = (A1MeasurementListCount > index ? "" : "");
            }
            else
            {
                text148.Text = (A1MeasurementListCount > index ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[index].Tube1_0).ToString() : "");
            }

            run148.Append(runProperties148);
            run148.Append(text148);

            paragraph98.Append(paragraphProperties98);
            paragraph98.Append(run148);

            tableCell86.Append(tableCellProperties86);
            tableCell86.Append(paragraph98);

            TableCell tableCell87 = new TableCell();

            TableCellProperties tableCellProperties87 = new TableCellProperties();
            TableCellWidth tableCellWidth87 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders87 = new TableCellBorders();
            TopBorder topBorder69 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder89 = new LeftBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder89 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders87.Append(topBorder69);
            tableCellBorders87.Append(leftBorder89);
            tableCellBorders87.Append(rightBorder89);
            Shading shading87 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties87.Append(tableCellWidth87);
            tableCellProperties87.Append(tableCellBorders87);
            tableCellProperties87.Append(shading87);

            Paragraph paragraph99 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties99 = new ParagraphProperties();
            Justification justification53 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties99 = new ParagraphMarkRunProperties();
            RunFonts runFonts248 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties99.Append(runFonts248);

            paragraphProperties99.Append(justification53);
            paragraphProperties99.Append(paragraphMarkRunProperties99);

            Run run151 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties151 = new RunProperties();
            RunFonts runFonts249 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize218 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript218 = new FontSizeComplexScript() { Val = "20" };

            runProperties151.Append(runFonts249);
            runProperties151.Append(fontSize218);
            runProperties151.Append(fontSizeComplexScript218);
            Text text151 = new Text();
            if (labSheetA1Sheet.LabSheetA1MeasurementList[index].Tube0_1 == null)
            {
                text151.Text = (A1MeasurementListCount > index ? "" : "");
            }
            else
            {
                text151.Text = (A1MeasurementListCount > index ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[index].Tube0_1).ToString() : "");
            }

            run151.Append(runProperties151);
            run151.Append(text151);

            paragraph99.Append(paragraphProperties99);
            paragraph99.Append(run151);

            tableCell87.Append(tableCellProperties87);
            tableCell87.Append(paragraph99);

            TableCell tableCell88 = new TableCell();

            TableCellProperties tableCellProperties88 = new TableCellProperties();
            TableCellWidth tableCellWidth88 = new TableCellWidth() { Width = "810", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders88 = new TableCellBorders();
            TopBorder topBorder70 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder90 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder90 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders88.Append(topBorder70);
            tableCellBorders88.Append(leftBorder90);
            tableCellBorders88.Append(rightBorder90);
            Shading shading88 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties88.Append(tableCellWidth88);
            tableCellProperties88.Append(tableCellBorders88);
            tableCellProperties88.Append(shading88);

            Paragraph paragraph100 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties100 = new ParagraphProperties();
            Justification justification54 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties100 = new ParagraphMarkRunProperties();
            RunFonts runFonts252 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties100.Append(runFonts252);

            paragraphProperties100.Append(justification54);
            paragraphProperties100.Append(paragraphMarkRunProperties100);

            Run run154 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties154 = new RunProperties();
            RunFonts runFonts253 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize221 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript221 = new FontSizeComplexScript() { Val = "20" };

            runProperties154.Append(runFonts253);
            runProperties154.Append(fontSize221);
            runProperties154.Append(fontSizeComplexScript221);
            Text text154 = new Text();

            if (labSheetA1Sheet.LabSheetA1MeasurementList[index].Salinity == null)
            {
                text154.Text = (A1MeasurementListCount > index ? "" : "");
            }
            else
            {
                text154.Text = (A1MeasurementListCount > index ? ((float)labSheetA1Sheet.LabSheetA1MeasurementList[index].Salinity).ToString("F1") : "");
            }

            run154.Append(runProperties154);
            run154.Append(text154);
            paragraph100.Append(paragraphProperties100);
            paragraph100.Append(run154);

            tableCell88.Append(tableCellProperties88);
            tableCell88.Append(paragraph100);

            TableCell tableCell89 = new TableCell();

            TableCellProperties tableCellProperties89 = new TableCellProperties();
            TableCellWidth tableCellWidth89 = new TableCellWidth() { Width = "810", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders89 = new TableCellBorders();
            TopBorder topBorder71 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder91 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder69 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder91 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders89.Append(topBorder71);
            tableCellBorders89.Append(leftBorder91);
            tableCellBorders89.Append(bottomBorder69);
            tableCellBorders89.Append(rightBorder91);
            Shading shading89 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties89.Append(tableCellWidth89);
            tableCellProperties89.Append(tableCellBorders89);
            tableCellProperties89.Append(shading89);

            Paragraph paragraph101 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties101 = new ParagraphProperties();
            Justification justification55 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties101 = new ParagraphMarkRunProperties();
            RunFonts runFonts257 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties101.Append(runFonts257);

            paragraphProperties101.Append(justification55);
            paragraphProperties101.Append(paragraphMarkRunProperties101);

            Run run158 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties158 = new RunProperties();
            RunFonts runFonts258 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize225 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript225 = new FontSizeComplexScript() { Val = "20" };

            runProperties158.Append(runFonts258);
            runProperties158.Append(fontSize225);
            runProperties158.Append(fontSizeComplexScript225);
            Text text158 = new Text();

            if (labSheetA1Sheet.LabSheetA1MeasurementList[index].Temperature == null)
            {
                text158.Text = (A1MeasurementListCount > index ? "" : "");
            }
            else
            {
                text158.Text = (A1MeasurementListCount > index ? ((float)labSheetA1Sheet.LabSheetA1MeasurementList[index].Temperature).ToString("F1") : "");
            }

            run158.Append(runProperties158);
            run158.Append(text158);

            paragraph101.Append(paragraphProperties101);
            paragraph101.Append(run158);

            tableCell89.Append(tableCellProperties89);
            tableCell89.Append(paragraph101);

            TableCell tableCell90 = new TableCell();

            TableCellProperties tableCellProperties90 = new TableCellProperties();
            TableCellWidth tableCellWidth90 = new TableCellWidth() { Width = "1024", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders90 = new TableCellBorders();
            TopBorder topBorder72 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder92 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder70 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder92 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders90.Append(topBorder72);
            tableCellBorders90.Append(leftBorder92);
            tableCellBorders90.Append(bottomBorder70);
            tableCellBorders90.Append(rightBorder92);
            Shading shading90 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties90.Append(tableCellWidth90);
            tableCellProperties90.Append(tableCellBorders90);
            tableCellProperties90.Append(shading90);

            Paragraph paragraph102 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties102 = new ParagraphProperties();
            Justification justification56 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties102 = new ParagraphMarkRunProperties();
            RunFonts runFonts262 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Highlight highlight2 = new Highlight() { Val = HighlightColorValues.Yellow };

            paragraphMarkRunProperties102.Append(runFonts262);
            paragraphMarkRunProperties102.Append(highlight2);

            paragraphProperties102.Append(justification56);
            paragraphProperties102.Append(paragraphMarkRunProperties102);

            Run run162 = new Run() { RsidRunProperties = "0072154A" };

            RunProperties runProperties162 = new RunProperties();
            RunFonts runFonts263 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize229 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript229 = new FontSizeComplexScript() { Val = "20" };

            runProperties162.Append(runFonts263);
            runProperties162.Append(fontSize229);
            runProperties162.Append(fontSizeComplexScript229);
            Text text162 = new Text();
            text162.Text = (A1MeasurementListCount > index ? labSheetA1Sheet.LabSheetA1MeasurementList[index].ProcessedBy : "");

            run162.Append(runProperties162);
            run162.Append(text162);

            paragraph102.Append(paragraphProperties102);
            paragraph102.Append(run162);

            tableCell90.Append(tableCellProperties90);
            tableCell90.Append(paragraph102);

            TableCell tableCell91 = new TableCell();

            TableCellProperties tableCellProperties91 = new TableCellProperties();
            TableCellWidth tableCellWidth91 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders91 = new TableCellBorders();
            TopBorder topBorder73 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder93 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder71 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder93 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders91.Append(topBorder73);
            tableCellBorders91.Append(leftBorder93);
            tableCellBorders91.Append(bottomBorder71);
            tableCellBorders91.Append(rightBorder93);
            Shading shading91 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties91.Append(tableCellWidth91);
            tableCellProperties91.Append(tableCellBorders91);
            tableCellProperties91.Append(shading91);

            Paragraph paragraph103 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

            ParagraphProperties paragraphProperties103 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties103 = new ParagraphMarkRunProperties();
            RunFonts runFonts267 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties103.Append(runFonts267);

            paragraphProperties103.Append(paragraphMarkRunProperties103);

            paragraph103.Append(paragraphProperties103);

            tableCell91.Append(tableCellProperties91);
            tableCell91.Append(paragraph103);

            if (index < 14 || (SheetNumber == 2 && (index - 34) < 14))
            {
                TableCell tableCell92 = new TableCell();

                TableCellProperties tableCellProperties92 = new TableCellProperties();
                TableCellWidth tableCellWidth92 = new TableCellWidth() { Width = "1008", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders92 = new TableCellBorders();
                TopBorder topBorder74 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder94 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder72 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder94 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders92.Append(topBorder74);
                tableCellBorders92.Append(leftBorder94);
                tableCellBorders92.Append(bottomBorder72);
                tableCellBorders92.Append(rightBorder94);
                Shading shading92 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties92.Append(tableCellWidth92);
                tableCellProperties92.Append(tableCellBorders92);
                tableCellProperties92.Append(shading92);

                Paragraph paragraph104 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties104 = new ParagraphProperties();
                Justification justification57 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties104 = new ParagraphMarkRunProperties();
                RunFonts runFonts268 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties104.Append(runFonts268);

                paragraphProperties104.Append(justification57);
                paragraphProperties104.Append(paragraphMarkRunProperties104);

                Run run166 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties166 = new RunProperties();
                RunFonts runFonts269 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize233 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript233 = new FontSizeComplexScript() { Val = "20" };

                runProperties166.Append(runFonts269);
                runProperties166.Append(fontSize233);
                runProperties166.Append(fontSizeComplexScript233);
                Text text166 = new Text();
                text166.Text = (A1MeasurementListCount > (index + 20) ? labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Site : "");

                run166.Append(runProperties166);
                run166.Append(text166);

                Run run167 = new Run();

                paragraph104.Append(paragraphProperties104);
                paragraph104.Append(run166);

                tableCell92.Append(tableCellProperties92);
                tableCell92.Append(paragraph104);

                TableCell tableCell93 = new TableCell();

                TableCellProperties tableCellProperties93 = new TableCellProperties();
                TableCellWidth tableCellWidth93 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders93 = new TableCellBorders();
                TopBorder topBorder75 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder95 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder73 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder95 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders93.Append(topBorder75);
                tableCellBorders93.Append(leftBorder95);
                tableCellBorders93.Append(bottomBorder73);
                tableCellBorders93.Append(rightBorder95);
                Shading shading93 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties93.Append(tableCellWidth93);
                tableCellProperties93.Append(tableCellBorders93);
                tableCellProperties93.Append(shading93);

                Paragraph paragraph105 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties105 = new ParagraphProperties();
                Justification justification58 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties105 = new ParagraphMarkRunProperties();
                RunFonts runFonts272 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties105.Append(runFonts272);

                paragraphProperties105.Append(justification58);
                paragraphProperties105.Append(paragraphMarkRunProperties105);

                Run run169 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties169 = new RunProperties();
                RunFonts runFonts273 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize236 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript236 = new FontSizeComplexScript() { Val = "20" };

                runProperties169.Append(runFonts273);
                runProperties169.Append(fontSize236);
                runProperties169.Append(fontSizeComplexScript236);
                Text text169 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Time == null)
                {
                    text169.Text = (A1MeasurementListCount > (index + 20) ? "" : "");
                }
                else
                {
                    text169.Text = (A1MeasurementListCount > (index + 20) ? ((DateTime)labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Time).ToString("hh:mm") : "");
                }

                run169.Append(runProperties169);
                run169.Append(text169);

                paragraph105.Append(paragraphProperties105);
                paragraph105.Append(run169);

                tableCell93.Append(tableCellProperties93);
                tableCell93.Append(paragraph105);

                TableCell tableCell94 = new TableCell();

                TableCellProperties tableCellProperties94 = new TableCellProperties();
                TableCellWidth tableCellWidth94 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders94 = new TableCellBorders();
                TopBorder topBorder76 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder96 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder74 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder96 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders94.Append(topBorder76);
                tableCellBorders94.Append(leftBorder96);
                tableCellBorders94.Append(bottomBorder74);
                tableCellBorders94.Append(rightBorder96);
                Shading shading94 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties94.Append(tableCellWidth94);
                tableCellProperties94.Append(tableCellBorders94);
                tableCellProperties94.Append(shading94);

                Paragraph paragraph106 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties106 = new ParagraphProperties();
                Justification justification59 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties106 = new ParagraphMarkRunProperties();
                RunFonts runFonts276 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties106.Append(runFonts276);

                paragraphProperties106.Append(justification59);
                paragraphProperties106.Append(paragraphMarkRunProperties106);

                Run run172 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties172 = new RunProperties();
                RunFonts runFonts277 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize239 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript239 = new FontSizeComplexScript() { Val = "20" };

                runProperties172.Append(runFonts277);
                runProperties172.Append(fontSize239);
                runProperties172.Append(fontSizeComplexScript239);
                Text text172 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].MPN == null)
                {
                    text172.Text = (A1MeasurementListCount > (index + 20) ? "" : "");
                }
                else
                {
                    text172.Text = (A1MeasurementListCount > (index + 20) ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].MPN).ToString() : "");
                }

                run172.Append(runProperties172);
                run172.Append(text172);

                paragraph106.Append(paragraphProperties106);
                paragraph106.Append(run172);

                tableCell94.Append(tableCellProperties94);
                tableCell94.Append(paragraph106);

                TableCell tableCell95 = new TableCell();

                TableCellProperties tableCellProperties95 = new TableCellProperties();
                TableCellWidth tableCellWidth95 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders95 = new TableCellBorders();
                TopBorder topBorder77 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder97 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder97 = new RightBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders95.Append(topBorder77);
                tableCellBorders95.Append(leftBorder97);
                tableCellBorders95.Append(rightBorder97);
                Shading shading95 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties95.Append(tableCellWidth95);
                tableCellProperties95.Append(tableCellBorders95);
                tableCellProperties95.Append(shading95);

                Paragraph paragraph107 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties107 = new ParagraphProperties();
                Justification justification60 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties107 = new ParagraphMarkRunProperties();
                RunFonts runFonts280 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties107.Append(runFonts280);

                paragraphProperties107.Append(justification60);
                paragraphProperties107.Append(paragraphMarkRunProperties107);

                Run run175 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties175 = new RunProperties();
                RunFonts runFonts281 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize242 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript242 = new FontSizeComplexScript() { Val = "20" };

                runProperties175.Append(runFonts281);
                runProperties175.Append(fontSize242);
                runProperties175.Append(fontSizeComplexScript242);
                Text text175 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Tube10 == null)
                {
                    text175.Text = (A1MeasurementListCount > (index + 20) ? "" : "");
                }
                else
                {
                    text175.Text = (A1MeasurementListCount > (index + 20) ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Tube10).ToString() : "");
                }

                run175.Append(runProperties175);
                run175.Append(text175);

                paragraph107.Append(paragraphProperties107);
                paragraph107.Append(run175);

                tableCell95.Append(tableCellProperties95);
                tableCell95.Append(paragraph107);

                TableCell tableCell96 = new TableCell();

                TableCellProperties tableCellProperties96 = new TableCellProperties();
                TableCellWidth tableCellWidth96 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders96 = new TableCellBorders();
                TopBorder topBorder78 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder98 = new LeftBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder98 = new RightBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders96.Append(topBorder78);
                tableCellBorders96.Append(leftBorder98);
                tableCellBorders96.Append(rightBorder98);
                Shading shading96 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties96.Append(tableCellWidth96);
                tableCellProperties96.Append(tableCellBorders96);
                tableCellProperties96.Append(shading96);

                Paragraph paragraph108 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties108 = new ParagraphProperties();
                Justification justification61 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties108 = new ParagraphMarkRunProperties();
                RunFonts runFonts284 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties108.Append(runFonts284);

                paragraphProperties108.Append(justification61);
                paragraphProperties108.Append(paragraphMarkRunProperties108);

                Run run178 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties178 = new RunProperties();
                RunFonts runFonts285 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize245 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript245 = new FontSizeComplexScript() { Val = "20" };

                runProperties178.Append(runFonts285);
                runProperties178.Append(fontSize245);
                runProperties178.Append(fontSizeComplexScript245);
                Text text178 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Tube1_0 == null)
                {
                    text178.Text = (A1MeasurementListCount > (index + 20) ? "" : "");
                }
                else
                {
                    text178.Text = (A1MeasurementListCount > (index + 20) ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Tube1_0).ToString() : "");
                }

                run178.Append(runProperties178);
                run178.Append(text178);

                paragraph108.Append(paragraphProperties108);
                paragraph108.Append(run178);

                tableCell96.Append(tableCellProperties96);
                tableCell96.Append(paragraph108);

                TableCell tableCell97 = new TableCell();

                TableCellProperties tableCellProperties97 = new TableCellProperties();
                TableCellWidth tableCellWidth97 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders97 = new TableCellBorders();
                TopBorder topBorder79 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder99 = new LeftBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder99 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders97.Append(topBorder79);
                tableCellBorders97.Append(leftBorder99);
                tableCellBorders97.Append(rightBorder99);
                Shading shading97 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties97.Append(tableCellWidth97);
                tableCellProperties97.Append(tableCellBorders97);
                tableCellProperties97.Append(shading97);

                Paragraph paragraph109 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties109 = new ParagraphProperties();
                Justification justification62 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties109 = new ParagraphMarkRunProperties();
                RunFonts runFonts288 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties109.Append(runFonts288);

                paragraphProperties109.Append(justification62);
                paragraphProperties109.Append(paragraphMarkRunProperties109);

                Run run181 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties181 = new RunProperties();
                RunFonts runFonts289 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize248 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript248 = new FontSizeComplexScript() { Val = "20" };

                runProperties181.Append(runFonts289);
                runProperties181.Append(fontSize248);
                runProperties181.Append(fontSizeComplexScript248);
                Text text181 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Tube0_1 == null)
                {
                    text181.Text = (A1MeasurementListCount > (index + 20) ? "" : "");
                }
                else
                {
                    text181.Text = (A1MeasurementListCount > (index + 20) ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Tube0_1).ToString() : "");
                }

                run181.Append(runProperties181);
                run181.Append(text181);

                paragraph109.Append(paragraphProperties109);
                paragraph109.Append(run181);

                tableCell97.Append(tableCellProperties97);
                tableCell97.Append(paragraph109);

                TableCell tableCell98 = new TableCell();

                TableCellProperties tableCellProperties98 = new TableCellProperties();
                TableCellWidth tableCellWidth98 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders98 = new TableCellBorders();
                TopBorder topBorder80 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder100 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder100 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders98.Append(topBorder80);
                tableCellBorders98.Append(leftBorder100);
                tableCellBorders98.Append(rightBorder100);
                Shading shading98 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties98.Append(tableCellWidth98);
                tableCellProperties98.Append(tableCellBorders98);
                tableCellProperties98.Append(shading98);

                Paragraph paragraph110 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties110 = new ParagraphProperties();
                Justification justification63 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties110 = new ParagraphMarkRunProperties();
                RunFonts runFonts292 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties110.Append(runFonts292);

                paragraphProperties110.Append(justification63);
                paragraphProperties110.Append(paragraphMarkRunProperties110);

                Run run184 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties184 = new RunProperties();
                RunFonts runFonts293 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize251 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript251 = new FontSizeComplexScript() { Val = "20" };

                runProperties184.Append(runFonts293);
                runProperties184.Append(fontSize251);
                runProperties184.Append(fontSizeComplexScript251);
                Text text184 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Salinity == null)
                {
                    text184.Text = (A1MeasurementListCount > (index + 20) ? "" : "");
                }
                else
                {
                    text184.Text = (A1MeasurementListCount > (index + 20) ? ((float)labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Salinity).ToString("F1") : "");
                }

                run184.Append(runProperties184);
                run184.Append(text184);

                paragraph110.Append(paragraphProperties110);
                paragraph110.Append(run184);

                tableCell98.Append(tableCellProperties98);
                tableCell98.Append(paragraph110);

                TableCell tableCell99 = new TableCell();

                TableCellProperties tableCellProperties99 = new TableCellProperties();
                TableCellWidth tableCellWidth99 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan5 = new GridSpan() { Val = 2 };

                TableCellBorders tableCellBorders99 = new TableCellBorders();
                TopBorder topBorder81 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder101 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder75 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder101 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders99.Append(topBorder81);
                tableCellBorders99.Append(leftBorder101);
                tableCellBorders99.Append(bottomBorder75);
                tableCellBorders99.Append(rightBorder101);
                Shading shading99 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties99.Append(tableCellWidth99);
                tableCellProperties99.Append(gridSpan5);
                tableCellProperties99.Append(tableCellBorders99);
                tableCellProperties99.Append(shading99);

                Paragraph paragraph111 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties111 = new ParagraphProperties();
                Justification justification64 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties111 = new ParagraphMarkRunProperties();
                RunFonts runFonts297 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties111.Append(runFonts297);

                paragraphProperties111.Append(justification64);
                paragraphProperties111.Append(paragraphMarkRunProperties111);

                Run run188 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties188 = new RunProperties();
                RunFonts runFonts298 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize255 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript255 = new FontSizeComplexScript() { Val = "20" };

                runProperties188.Append(runFonts298);
                runProperties188.Append(fontSize255);
                runProperties188.Append(fontSizeComplexScript255);
                Text text188 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Temperature == null)
                {
                    text188.Text = (A1MeasurementListCount > (index + 20) ? "" : "");
                }
                else
                {
                    text188.Text = (A1MeasurementListCount > (index + 20) ? ((float)labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].Temperature).ToString("F1") : "");
                }

                run188.Append(runProperties188);
                run188.Append(text188);

                paragraph111.Append(paragraphProperties111);
                paragraph111.Append(run188);

                tableCell99.Append(tableCellProperties99);
                tableCell99.Append(paragraph111);

                TableCell tableCell100 = new TableCell();

                TableCellProperties tableCellProperties100 = new TableCellProperties();
                TableCellWidth tableCellWidth100 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders100 = new TableCellBorders();
                TopBorder topBorder82 = new TopBorder() { Val = topBorderValues, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder102 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder76 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder102 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders100.Append(topBorder82);
                tableCellBorders100.Append(leftBorder102);
                tableCellBorders100.Append(bottomBorder76);
                tableCellBorders100.Append(rightBorder102);
                Shading shading100 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties100.Append(tableCellWidth100);
                tableCellProperties100.Append(tableCellBorders100);
                tableCellProperties100.Append(shading100);

                Paragraph paragraph112 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties112 = new ParagraphProperties();
                Justification justification65 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties112 = new ParagraphMarkRunProperties();
                RunFonts runFonts302 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Highlight highlight3 = new Highlight() { Val = HighlightColorValues.Yellow };

                paragraphMarkRunProperties112.Append(runFonts302);
                paragraphMarkRunProperties112.Append(highlight3);

                paragraphProperties112.Append(justification65);
                paragraphProperties112.Append(paragraphMarkRunProperties112);

                Run run192 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties192 = new RunProperties();
                RunFonts runFonts303 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize259 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript259 = new FontSizeComplexScript() { Val = "20" };

                runProperties192.Append(runFonts303);
                runProperties192.Append(fontSize259);
                runProperties192.Append(fontSizeComplexScript259);
                Text text192 = new Text();
                text192.Text = (A1MeasurementListCount > (index + 20) ? labSheetA1Sheet.LabSheetA1MeasurementList[index + 20].ProcessedBy : "");

                run192.Append(runProperties192);
                run192.Append(text192);
                paragraph112.Append(paragraphProperties112);
                paragraph112.Append(run192);

                tableCell100.Append(tableCellProperties100);
                tableCell100.Append(paragraph112);

                tableRow.Append(tableRowProperties4);
                tableRow.Append(tableCell82);
                tableRow.Append(tableCell83);
                tableRow.Append(tableCell84);
                tableRow.Append(tableCell85);
                tableRow.Append(tableCell86);
                tableRow.Append(tableCell87);
                tableRow.Append(tableCell88);
                tableRow.Append(tableCell89);
                tableRow.Append(tableCell90);
                tableRow.Append(tableCell91);
                tableRow.Append(tableCell92);
                tableRow.Append(tableCell93);
                tableRow.Append(tableCell94);
                tableRow.Append(tableCell95);
                tableRow.Append(tableCell96);
                tableRow.Append(tableCell97);
                tableRow.Append(tableCell98);
                tableRow.Append(tableCell99);
                tableRow.Append(tableCell100);
            }
            else if (index == 14 || (SheetNumber == 2 && (index - 34) == 14))
            {
                TableCell tableCell358 = new TableCell();

                TableCellProperties tableCellProperties358 = new TableCellProperties();
                TableCellWidth tableCellWidth358 = new TableCellWidth() { Width = "7668", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan19 = new GridSpan() { Val = 10 };

                TableCellBorders tableCellBorders358 = new TableCellBorders();
                TopBorder topBorder248 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder360 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder270 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder360 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders358.Append(topBorder248);
                tableCellBorders358.Append(leftBorder360);
                tableCellBorders358.Append(bottomBorder270);
                tableCellBorders358.Append(rightBorder360);
                Shading shading358 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                tableCellProperties358.Append(tableCellWidth358);
                tableCellProperties358.Append(gridSpan19);
                tableCellProperties358.Append(tableCellBorders358);
                tableCellProperties358.Append(shading358);

                Paragraph paragraph370 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00002800", RsidParagraphProperties = "00002800", RsidRunAdditionDefault = "00002800" };

                ParagraphProperties paragraphProperties370 = new ParagraphProperties();
                Justification justification73 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties370 = new ParagraphMarkRunProperties();
                RunFonts runFonts560 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties370.Append(runFonts560);

                paragraphProperties370.Append(justification73);
                paragraphProperties370.Append(paragraphMarkRunProperties370);

                Run run192 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties192 = new RunProperties();
                RunFonts runFonts561 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize259 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript259 = new FontSizeComplexScript() { Val = "20" };

                runProperties192.Append(runFonts561);
                runProperties192.Append(fontSize259);
                runProperties192.Append(fontSizeComplexScript259);
                Text text192 = new Text();
                text192.Text = "Sheet " + SheetNumber + " of " + (A1MeasurementListCount > 34 ? "2" : "1");

                run192.Append(runProperties192);
                run192.Append(text192);
                paragraph370.Append(paragraphProperties370);
                paragraph370.Append(run192);

                tableCell358.Append(tableCellProperties358);
                tableCell358.Append(paragraph370);

                tableRow.Append(tableRowProperties4);
                tableRow.Append(tableCell82);
                tableRow.Append(tableCell83);
                tableRow.Append(tableCell84);
                tableRow.Append(tableCell85);
                tableRow.Append(tableCell86);
                tableRow.Append(tableCell87);
                tableRow.Append(tableCell88);
                tableRow.Append(tableCell89);
                tableRow.Append(tableCell90);
                tableRow.Append(tableCell91);
                tableRow.Append(tableCell358);
            }
            else if (index == 15 || (SheetNumber == 2 && (index - 34) == 15))
            {
                // Create Dupliate Row Part

                TableCell tableCell377 = new TableCell();

                TableCellProperties tableCellProperties377 = new TableCellProperties();
                TableCellWidth tableCellWidth377 = new TableCellWidth() { Width = "1008", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge23 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders377 = new TableCellBorders();
                TopBorder topBorder263 = new TopBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder379 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                RightBorder rightBorder379 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders377.Append(topBorder263);
                tableCellBorders377.Append(leftBorder379);
                tableCellBorders377.Append(rightBorder379);
                Shading shading377 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties377.Append(tableCellWidth377);
                tableCellProperties377.Append(verticalMerge23);
                tableCellProperties377.Append(tableCellBorders377);
                tableCellProperties377.Append(shading377);

                Paragraph paragraph389 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties389 = new ParagraphProperties();
                Justification justification71 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties389 = new ParagraphMarkRunProperties();
                RunFonts runFonts599 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color8 = new Color() { Val = "808080" };
                FontSize fontSize279 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript279 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties389.Append(runFonts599);
                paragraphMarkRunProperties389.Append(color8);
                paragraphMarkRunProperties389.Append(fontSize279);
                paragraphMarkRunProperties389.Append(fontSizeComplexScript279);

                paragraphProperties389.Append(justification71);
                paragraphProperties389.Append(paragraphMarkRunProperties389);

                Run run212 = new Run() { RsidRunProperties = "00ED7FB2" };

                RunProperties runProperties212 = new RunProperties();
                RunFonts runFonts600 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color9 = new Color() { Val = "808080" };
                FontSize fontSize280 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript280 = new FontSizeComplexScript() { Val = "16" };

                runProperties212.Append(runFonts600);
                runProperties212.Append(color9);
                runProperties212.Append(fontSize280);
                runProperties212.Append(fontSizeComplexScript280);
                Text text212 = new Text();
                text212.Text = "Duplicate #";

                run212.Append(runProperties212);
                run212.Append(text212);

                paragraph389.Append(paragraphProperties389);
                paragraph389.Append(run212);

                Paragraph paragraph390 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties390 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties390 = new ParagraphMarkRunProperties();
                RunFonts runFonts601 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color10 = new Color() { Val = "808080" };
                FontSize fontSize281 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript281 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties390.Append(runFonts601);
                paragraphMarkRunProperties390.Append(color10);
                paragraphMarkRunProperties390.Append(fontSize281);
                paragraphMarkRunProperties390.Append(fontSizeComplexScript281);

                paragraphProperties390.Append(paragraphMarkRunProperties390);

                paragraph390.Append(paragraphProperties390);

                Paragraph paragraph391 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties391 = new ParagraphProperties();
                Justification justification72 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties391 = new ParagraphMarkRunProperties();
                RunFonts runFonts602 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color11 = new Color() { Val = "808080" };
                FontSize fontSize282 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript282 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties391.Append(runFonts602);
                paragraphMarkRunProperties391.Append(color11);
                paragraphMarkRunProperties391.Append(fontSize282);
                paragraphMarkRunProperties391.Append(fontSizeComplexScript282);

                paragraphProperties391.Append(justification72);
                paragraphProperties391.Append(paragraphMarkRunProperties391);

                Run run213 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties213 = new RunProperties();
                RunFonts runFonts603 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize283 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript283 = new FontSizeComplexScript() { Val = "20" };

                runProperties213.Append(runFonts603);
                runProperties213.Append(fontSize283);
                runProperties213.Append(fontSizeComplexScript283);
                Text text213 = new Text();
                text213.Text = (DuplicateIndex > 0 ? labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].Site : "");

                run213.Append(runProperties213);
                run213.Append(text213);

                paragraph391.Append(paragraphProperties391);
                paragraph391.Append(run213);

                tableCell377.Append(tableCellProperties377);
                tableCell377.Append(paragraph389);
                tableCell377.Append(paragraph390);
                tableCell377.Append(paragraph391);

                TableCell tableCell378 = new TableCell();

                TableCellProperties tableCellProperties378 = new TableCellProperties();
                TableCellWidth tableCellWidth378 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge24 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders378 = new TableCellBorders();
                TopBorder topBorder264 = new TopBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder380 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder380 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders378.Append(topBorder264);
                tableCellBorders378.Append(leftBorder380);
                tableCellBorders378.Append(rightBorder380);
                Shading shading378 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties378.Append(tableCellWidth378);
                tableCellProperties378.Append(verticalMerge24);
                tableCellProperties378.Append(tableCellBorders378);
                tableCellProperties378.Append(shading378);

                Paragraph paragraph392 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties392 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties392 = new ParagraphMarkRunProperties();
                RunFonts runFonts606 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color12 = new Color() { Val = "808080" };
                FontSize fontSize286 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript286 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties392.Append(runFonts606);
                paragraphMarkRunProperties392.Append(color12);
                paragraphMarkRunProperties392.Append(fontSize286);
                paragraphMarkRunProperties392.Append(fontSizeComplexScript286);

                paragraphProperties392.Append(paragraphMarkRunProperties392);

                Run run216 = new Run() { RsidRunProperties = "00ED7FB2" };

                RunProperties runProperties216 = new RunProperties();
                RunFonts runFonts607 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color13 = new Color() { Val = "808080" };
                FontSize fontSize287 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript287 = new FontSizeComplexScript() { Val = "16" };

                runProperties216.Append(runFonts607);
                runProperties216.Append(color13);
                runProperties216.Append(fontSize287);
                runProperties216.Append(fontSizeComplexScript287);
                Text text216 = new Text();
                text216.Text = "Analyst";

                run216.Append(runProperties216);
                run216.Append(text216);

                paragraph392.Append(paragraphProperties392);
                paragraph392.Append(run216);

                Paragraph paragraph393 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties393 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties393 = new ParagraphMarkRunProperties();
                RunFonts runFonts608 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color14 = new Color() { Val = "808080" };
                FontSize fontSize288 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript288 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties393.Append(runFonts608);
                paragraphMarkRunProperties393.Append(color14);
                paragraphMarkRunProperties393.Append(fontSize288);
                paragraphMarkRunProperties393.Append(fontSizeComplexScript288);

                paragraphProperties393.Append(paragraphMarkRunProperties393);

                paragraph393.Append(paragraphProperties393);

                Paragraph paragraph394 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties394 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties394 = new ParagraphMarkRunProperties();
                RunFonts runFonts609 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color15 = new Color() { Val = "808080" };
                FontSize fontSize289 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript289 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties394.Append(runFonts609);
                paragraphMarkRunProperties394.Append(color15);
                paragraphMarkRunProperties394.Append(fontSize289);
                paragraphMarkRunProperties394.Append(fontSizeComplexScript289);

                paragraphProperties394.Append(paragraphMarkRunProperties394);

                paragraph394.Append(paragraphProperties394);

                Paragraph paragraph395 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties395 = new ParagraphProperties();
                Justification justification73 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties395 = new ParagraphMarkRunProperties();
                RunFonts runFonts610 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize290 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript290 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties395.Append(runFonts610);
                paragraphMarkRunProperties395.Append(fontSize290);
                paragraphMarkRunProperties395.Append(fontSizeComplexScript290);

                paragraphProperties395.Append(justification73);
                paragraphProperties395.Append(paragraphMarkRunProperties395);

                Run run217 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties217 = new RunProperties();
                RunFonts runFonts611 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize291 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript291 = new FontSizeComplexScript() { Val = "20" };

                runProperties217.Append(runFonts611);
                runProperties217.Append(fontSize291);
                runProperties217.Append(fontSizeComplexScript291);
                Text text217 = new Text();
                text217.Text = (DuplicateIndex > 0 ? labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].ProcessedBy : "");

                run217.Append(runProperties217);
                run217.Append(text217);

                paragraph395.Append(paragraphProperties395);
                paragraph395.Append(run217);

                tableCell378.Append(tableCellProperties378);
                tableCell378.Append(paragraph392);
                tableCell378.Append(paragraph393);
                tableCell378.Append(paragraph394);
                tableCell378.Append(paragraph395);

                TableCell tableCell379 = new TableCell();

                TableCellProperties tableCellProperties379 = new TableCellProperties();
                TableCellWidth tableCellWidth379 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge25 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders379 = new TableCellBorders();
                TopBorder topBorder265 = new TopBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder381 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder381 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders379.Append(topBorder265);
                tableCellBorders379.Append(leftBorder381);
                tableCellBorders379.Append(rightBorder381);
                Shading shading379 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties379.Append(tableCellWidth379);
                tableCellProperties379.Append(verticalMerge25);
                tableCellProperties379.Append(tableCellBorders379);
                tableCellProperties379.Append(shading379);

                Paragraph paragraph396 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties396 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties396 = new ParagraphMarkRunProperties();
                RunFonts runFonts614 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize294 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript294 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties396.Append(runFonts614);
                paragraphMarkRunProperties396.Append(fontSize294);
                paragraphMarkRunProperties396.Append(fontSizeComplexScript294);

                paragraphProperties396.Append(paragraphMarkRunProperties396);

                paragraph396.Append(paragraphProperties396);

                Paragraph paragraph397 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties397 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties397 = new ParagraphMarkRunProperties();
                RunFonts runFonts615 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize295 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript295 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties397.Append(runFonts615);
                paragraphMarkRunProperties397.Append(fontSize295);
                paragraphMarkRunProperties397.Append(fontSizeComplexScript295);

                paragraphProperties397.Append(paragraphMarkRunProperties397);

                paragraph397.Append(paragraphProperties397);

                Paragraph paragraph398 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties398 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties398 = new ParagraphMarkRunProperties();
                RunFonts runFonts616 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize296 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript296 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties398.Append(runFonts616);
                paragraphMarkRunProperties398.Append(fontSize296);
                paragraphMarkRunProperties398.Append(fontSizeComplexScript296);

                paragraphProperties398.Append(paragraphMarkRunProperties398);

                paragraph398.Append(paragraphProperties398);

                Paragraph paragraph399 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties399 = new ParagraphProperties();
                Justification justification74 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties399 = new ParagraphMarkRunProperties();
                RunFonts runFonts617 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize297 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript297 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties399.Append(runFonts617);
                paragraphMarkRunProperties399.Append(fontSize297);
                paragraphMarkRunProperties399.Append(fontSizeComplexScript297);

                paragraphProperties399.Append(justification74);
                paragraphProperties399.Append(paragraphMarkRunProperties399);

                Run run220 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties220 = new RunProperties();
                RunFonts runFonts618 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize298 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript298 = new FontSizeComplexScript() { Val = "20" };

                runProperties220.Append(runFonts618);
                runProperties220.Append(fontSize298);
                runProperties220.Append(fontSizeComplexScript298);
                Text text220 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].MPN == null)
                {
                    text220.Text = (DuplicateIndex > 0 ? "" : "");
                }
                else
                {
                    text220.Text = (DuplicateIndex > 0 ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].MPN).ToString() : "");
                }

                run220.Append(runProperties220);
                run220.Append(text220);

                paragraph399.Append(paragraphProperties399);
                paragraph399.Append(run220);

                tableCell379.Append(tableCellProperties379);
                tableCell379.Append(paragraph396);
                tableCell379.Append(paragraph397);
                tableCell379.Append(paragraph398);
                tableCell379.Append(paragraph399);

                TableCell tableCell380 = new TableCell();

                TableCellProperties tableCellProperties380 = new TableCellProperties();
                TableCellWidth tableCellWidth380 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge26 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders380 = new TableCellBorders();
                TopBorder topBorder266 = new TopBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder382 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder382 = new RightBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders380.Append(topBorder266);
                tableCellBorders380.Append(leftBorder382);
                tableCellBorders380.Append(rightBorder382);
                Shading shading380 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties380.Append(tableCellWidth380);
                tableCellProperties380.Append(verticalMerge26);
                tableCellProperties380.Append(tableCellBorders380);
                tableCellProperties380.Append(shading380);

                Paragraph paragraph400 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties400 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties400 = new ParagraphMarkRunProperties();
                RunFonts runFonts621 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize301 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript301 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties400.Append(runFonts621);
                paragraphMarkRunProperties400.Append(fontSize301);
                paragraphMarkRunProperties400.Append(fontSizeComplexScript301);

                paragraphProperties400.Append(paragraphMarkRunProperties400);

                paragraph400.Append(paragraphProperties400);

                Paragraph paragraph401 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties401 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties401 = new ParagraphMarkRunProperties();
                RunFonts runFonts622 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize302 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript302 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties401.Append(runFonts622);
                paragraphMarkRunProperties401.Append(fontSize302);
                paragraphMarkRunProperties401.Append(fontSizeComplexScript302);

                paragraphProperties401.Append(paragraphMarkRunProperties401);

                paragraph401.Append(paragraphProperties401);

                Paragraph paragraph402 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties402 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties402 = new ParagraphMarkRunProperties();
                RunFonts runFonts623 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize303 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript303 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties402.Append(runFonts623);
                paragraphMarkRunProperties402.Append(fontSize303);
                paragraphMarkRunProperties402.Append(fontSizeComplexScript303);

                paragraphProperties402.Append(paragraphMarkRunProperties402);

                paragraph402.Append(paragraphProperties402);

                Paragraph paragraph403 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties403 = new ParagraphProperties();
                Justification justification75 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties403 = new ParagraphMarkRunProperties();
                RunFonts runFonts624 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize304 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript304 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties403.Append(runFonts624);
                paragraphMarkRunProperties403.Append(fontSize304);
                paragraphMarkRunProperties403.Append(fontSizeComplexScript304);

                paragraphProperties403.Append(justification75);
                paragraphProperties403.Append(paragraphMarkRunProperties403);

                Run run223 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties223 = new RunProperties();
                RunFonts runFonts625 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize305 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript305 = new FontSizeComplexScript() { Val = "20" };

                runProperties223.Append(runFonts625);
                runProperties223.Append(fontSize305);
                runProperties223.Append(fontSizeComplexScript305);
                Text text223 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].Tube10 == null)
                {
                    text223.Text = (DuplicateIndex > 0 ? "" : "");
                }
                else
                {
                    text223.Text = (DuplicateIndex > 0 ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].Tube10).ToString() : "");
                }

                run223.Append(runProperties223);
                run223.Append(text223);

                paragraph403.Append(paragraphProperties403);
                paragraph403.Append(run223);

                tableCell380.Append(tableCellProperties380);
                tableCell380.Append(paragraph400);
                tableCell380.Append(paragraph401);
                tableCell380.Append(paragraph402);
                tableCell380.Append(paragraph403);

                TableCell tableCell381 = new TableCell();

                TableCellProperties tableCellProperties381 = new TableCellProperties();
                TableCellWidth tableCellWidth381 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge27 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders381 = new TableCellBorders();
                TopBorder topBorder267 = new TopBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder383 = new LeftBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder383 = new RightBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders381.Append(topBorder267);
                tableCellBorders381.Append(leftBorder383);
                tableCellBorders381.Append(rightBorder383);
                Shading shading381 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties381.Append(tableCellWidth381);
                tableCellProperties381.Append(verticalMerge27);
                tableCellProperties381.Append(tableCellBorders381);
                tableCellProperties381.Append(shading381);

                Paragraph paragraph404 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties404 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties404 = new ParagraphMarkRunProperties();
                RunFonts runFonts628 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize308 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript308 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties404.Append(runFonts628);
                paragraphMarkRunProperties404.Append(fontSize308);
                paragraphMarkRunProperties404.Append(fontSizeComplexScript308);

                paragraphProperties404.Append(paragraphMarkRunProperties404);

                paragraph404.Append(paragraphProperties404);

                Paragraph paragraph405 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties405 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties405 = new ParagraphMarkRunProperties();
                RunFonts runFonts629 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize309 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript309 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties405.Append(runFonts629);
                paragraphMarkRunProperties405.Append(fontSize309);
                paragraphMarkRunProperties405.Append(fontSizeComplexScript309);

                paragraphProperties405.Append(paragraphMarkRunProperties405);

                paragraph405.Append(paragraphProperties405);

                Paragraph paragraph406 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties406 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties406 = new ParagraphMarkRunProperties();
                RunFonts runFonts630 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize310 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript310 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties406.Append(runFonts630);
                paragraphMarkRunProperties406.Append(fontSize310);
                paragraphMarkRunProperties406.Append(fontSizeComplexScript310);

                paragraphProperties406.Append(paragraphMarkRunProperties406);

                paragraph406.Append(paragraphProperties406);

                Paragraph paragraph407 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties407 = new ParagraphProperties();
                Justification justification76 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties407 = new ParagraphMarkRunProperties();
                RunFonts runFonts631 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize311 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript311 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties407.Append(runFonts631);
                paragraphMarkRunProperties407.Append(fontSize311);
                paragraphMarkRunProperties407.Append(fontSizeComplexScript311);

                paragraphProperties407.Append(justification76);
                paragraphProperties407.Append(paragraphMarkRunProperties407);

                Run run226 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties226 = new RunProperties();
                RunFonts runFonts632 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize312 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript312 = new FontSizeComplexScript() { Val = "20" };

                runProperties226.Append(runFonts632);
                runProperties226.Append(fontSize312);
                runProperties226.Append(fontSizeComplexScript312);
                Text text226 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].Tube1_0 == null)
                {
                    text226.Text = (DuplicateIndex > 0 ? "" : "");
                }
                else
                {
                    text226.Text = (DuplicateIndex > 0 ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].Tube1_0).ToString() : "");
                }

                run226.Append(runProperties226);
                run226.Append(text226);

                paragraph407.Append(paragraphProperties407);
                paragraph407.Append(run226);

                tableCell381.Append(tableCellProperties381);
                tableCell381.Append(paragraph404);
                tableCell381.Append(paragraph405);
                tableCell381.Append(paragraph406);
                tableCell381.Append(paragraph407);

                TableCell tableCell382 = new TableCell();

                TableCellProperties tableCellProperties382 = new TableCellProperties();
                TableCellWidth tableCellWidth382 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge28 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders382 = new TableCellBorders();
                TopBorder topBorder268 = new TopBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder384 = new LeftBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder384 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders382.Append(topBorder268);
                tableCellBorders382.Append(leftBorder384);
                tableCellBorders382.Append(rightBorder384);
                Shading shading382 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties382.Append(tableCellWidth382);
                tableCellProperties382.Append(verticalMerge28);
                tableCellProperties382.Append(tableCellBorders382);
                tableCellProperties382.Append(shading382);

                Paragraph paragraph408 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties408 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties408 = new ParagraphMarkRunProperties();
                RunFonts runFonts635 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize315 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript315 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties408.Append(runFonts635);
                paragraphMarkRunProperties408.Append(fontSize315);
                paragraphMarkRunProperties408.Append(fontSizeComplexScript315);

                paragraphProperties408.Append(paragraphMarkRunProperties408);

                paragraph408.Append(paragraphProperties408);

                Paragraph paragraph409 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties409 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties409 = new ParagraphMarkRunProperties();
                RunFonts runFonts636 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize316 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript316 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties409.Append(runFonts636);
                paragraphMarkRunProperties409.Append(fontSize316);
                paragraphMarkRunProperties409.Append(fontSizeComplexScript316);

                paragraphProperties409.Append(paragraphMarkRunProperties409);

                paragraph409.Append(paragraphProperties409);

                Paragraph paragraph410 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties410 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties410 = new ParagraphMarkRunProperties();
                RunFonts runFonts637 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize317 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript317 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties410.Append(runFonts637);
                paragraphMarkRunProperties410.Append(fontSize317);
                paragraphMarkRunProperties410.Append(fontSizeComplexScript317);

                paragraphProperties410.Append(paragraphMarkRunProperties410);

                paragraph410.Append(paragraphProperties410);

                Paragraph paragraph411 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties411 = new ParagraphProperties();
                Justification justification77 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties411 = new ParagraphMarkRunProperties();
                RunFonts runFonts638 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize318 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript318 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties411.Append(runFonts638);
                paragraphMarkRunProperties411.Append(fontSize318);
                paragraphMarkRunProperties411.Append(fontSizeComplexScript318);

                paragraphProperties411.Append(justification77);
                paragraphProperties411.Append(paragraphMarkRunProperties411);

                Run run229 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties229 = new RunProperties();
                RunFonts runFonts639 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize319 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript319 = new FontSizeComplexScript() { Val = "20" };

                runProperties229.Append(runFonts639);
                runProperties229.Append(fontSize319);
                runProperties229.Append(fontSizeComplexScript319);
                Text text229 = new Text();

                if (labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].Tube0_1 == null)
                {
                    text229.Text = (DuplicateIndex > 0 ? "" : "");
                }
                else
                {
                    text229.Text = (DuplicateIndex > 0 ? ((int)labSheetA1Sheet.LabSheetA1MeasurementList[DuplicateIndex].Tube0_1).ToString() : "");
                }

                run229.Append(runProperties229);
                run229.Append(text229);

                paragraph411.Append(paragraphProperties411);
                paragraph411.Append(run229);

                tableCell382.Append(tableCellProperties382);
                tableCell382.Append(paragraph408);
                tableCell382.Append(paragraph409);
                tableCell382.Append(paragraph410);
                tableCell382.Append(paragraph411);

                TableCell tableCell383 = new TableCell();

                TableCellProperties tableCellProperties383 = new TableCellProperties();
                TableCellWidth tableCellWidth383 = new TableCellWidth() { Width = "2700", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan20 = new GridSpan() { Val = 4 };
                VerticalMerge verticalMerge29 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders383 = new TableCellBorders();
                TopBorder topBorder269 = new TopBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder385 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder385 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders383.Append(topBorder269);
                tableCellBorders383.Append(leftBorder385);
                tableCellBorders383.Append(rightBorder385);
                Shading shading383 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties383.Append(tableCellWidth383);
                tableCellProperties383.Append(gridSpan20);
                tableCellProperties383.Append(verticalMerge29);
                tableCellProperties383.Append(tableCellBorders383);
                tableCellProperties383.Append(shading383);

                Paragraph paragraph412 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties412 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties412 = new ParagraphMarkRunProperties();
                RunFonts runFonts642 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color16 = new Color() { Val = "808080" };
                FontSize fontSize322 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript322 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties412.Append(runFonts642);
                paragraphMarkRunProperties412.Append(color16);
                paragraphMarkRunProperties412.Append(fontSize322);
                paragraphMarkRunProperties412.Append(fontSizeComplexScript322);

                paragraphProperties412.Append(paragraphMarkRunProperties412);

                Run run232 = new Run() { RsidRunProperties = "00ED7FB2" };

                RunProperties runProperties232 = new RunProperties();
                RunFonts runFonts643 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color17 = new Color() { Val = "808080" };
                FontSize fontSize323 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript323 = new FontSizeComplexScript() { Val = "16" };

                runProperties232.Append(runFonts643);
                runProperties232.Append(color17);
                runProperties232.Append(fontSize323);
                runProperties232.Append(fontSizeComplexScript323);
                Text text232 = new Text();
                text232.Text = "Data Entry Date/ Initials";

                run232.Append(runProperties232);
                run232.Append(text232);

                paragraph412.Append(paragraphProperties412);
                paragraph412.Append(run232);

                Paragraph paragraph413 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties413 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties413 = new ParagraphMarkRunProperties();
                RunFonts runFonts644 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color18 = new Color() { Val = "808080" };
                FontSize fontSize324 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript324 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties413.Append(runFonts644);
                paragraphMarkRunProperties413.Append(color18);
                paragraphMarkRunProperties413.Append(fontSize324);
                paragraphMarkRunProperties413.Append(fontSizeComplexScript324);

                paragraphProperties413.Append(paragraphMarkRunProperties413);

                paragraph413.Append(paragraphProperties413);

                Paragraph paragraph414 = new Paragraph() { RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties414 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties414 = new ParagraphMarkRunProperties();
                RunFonts runFonts645 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color19 = new Color() { Val = "808080" };
                FontSize fontSize325 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript325 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties414.Append(runFonts645);
                paragraphMarkRunProperties414.Append(color19);
                paragraphMarkRunProperties414.Append(fontSize325);
                paragraphMarkRunProperties414.Append(fontSizeComplexScript325);

                paragraphProperties414.Append(paragraphMarkRunProperties414);

                paragraph414.Append(paragraphProperties414);

                Paragraph paragraph415 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00FB2E61", RsidParagraphProperties = "00FB2E61", RsidRunAdditionDefault = "00FB2E61" };

                ParagraphProperties paragraphProperties415 = new ParagraphProperties();
                Justification justification78 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties415 = new ParagraphMarkRunProperties();
                RunFonts runFonts646 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color20 = new Color() { Val = "808080" };
                FontSize fontSize326 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript326 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties415.Append(runFonts646);
                paragraphMarkRunProperties415.Append(color20);
                paragraphMarkRunProperties415.Append(fontSize326);
                paragraphMarkRunProperties415.Append(fontSizeComplexScript326);

                paragraphProperties415.Append(justification78);
                paragraphProperties415.Append(paragraphMarkRunProperties415);

                Run run233 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties233 = new RunProperties();
                RunFonts runFonts647 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize327 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript327 = new FontSizeComplexScript() { Val = "20" };

                runProperties233.Append(runFonts647);
                runProperties233.Append(fontSize327);
                runProperties233.Append(fontSizeComplexScript327);
                Text text233 = new Text();
                DateTime DateTimeDuplicateDataEntryDate = new DateTime(int.Parse(labSheetA1Sheet.DuplicateDataEntryYear), int.Parse(labSheetA1Sheet.DuplicateDataEntryMonth), int.Parse(labSheetA1Sheet.DuplicateDataEntryDay));
                text233.Text = DateTimeDuplicateDataEntryDate.ToString("yyyy MMMM dd") + "      " + labSheetA1Sheet.DuplicateDataEntryInitials;

                run233.Append(runProperties233);
                run233.Append(text233);

                paragraph415.Append(paragraphProperties415);
                paragraph415.Append(run233);

                tableCell383.Append(tableCellProperties383);
                tableCell383.Append(paragraph412);
                tableCell383.Append(paragraph413);
                tableCell383.Append(paragraph414);
                tableCell383.Append(paragraph415);

                tableRow.Append(tableRowProperties4);
                tableRow.Append(tableCell82);
                tableRow.Append(tableCell83);
                tableRow.Append(tableCell84);
                tableRow.Append(tableCell85);
                tableRow.Append(tableCell86);
                tableRow.Append(tableCell87);
                tableRow.Append(tableCell88);
                tableRow.Append(tableCell89);
                tableRow.Append(tableCell90);
                tableRow.Append(tableCell91);
                tableRow.Append(tableCell377);
                tableRow.Append(tableCell378);
                tableRow.Append(tableCell379);
                tableRow.Append(tableCell380);
                tableRow.Append(tableCell381);
                tableRow.Append(tableCell382);
                tableRow.Append(tableCell383);
            }
            else if (index == 16 || index == 17 || (SheetNumber == 2 && (index - 34) == 16) || (SheetNumber == 2 && (index - 34) == 17))
            {
                TableCell tableCell411 = new TableCell();

                TableCellProperties tableCellProperties411 = new TableCellProperties();
                TableCellWidth tableCellWidth411 = new TableCellWidth() { Width = "1008", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge37 = new VerticalMerge();

                TableCellBorders tableCellBorders411 = new TableCellBorders();
                TopBorder topBorder285 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder413 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder297 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder413 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders411.Append(topBorder285);
                tableCellBorders411.Append(leftBorder413);
                tableCellBorders411.Append(bottomBorder297);
                tableCellBorders411.Append(rightBorder413);
                Shading shading411 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties411.Append(tableCellWidth411);
                tableCellProperties411.Append(verticalMerge37);
                tableCellProperties411.Append(tableCellBorders411);
                tableCellProperties411.Append(shading411);

                Paragraph paragraph443 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties443 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties443 = new ParagraphMarkRunProperties();
                RunFonts runFonts728 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color21 = new Color() { Val = "808080" };
                FontSize fontSize417 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript417 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties443.Append(runFonts728);
                paragraphMarkRunProperties443.Append(color21);
                paragraphMarkRunProperties443.Append(fontSize417);
                paragraphMarkRunProperties443.Append(fontSizeComplexScript417);

                paragraphProperties443.Append(paragraphMarkRunProperties443);

                paragraph443.Append(paragraphProperties443);

                tableCell411.Append(tableCellProperties411);
                tableCell411.Append(paragraph443);

                TableCell tableCell412 = new TableCell();

                TableCellProperties tableCellProperties412 = new TableCellProperties();
                TableCellWidth tableCellWidth412 = new TableCellWidth() { Width = "900", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge38 = new VerticalMerge();

                TableCellBorders tableCellBorders412 = new TableCellBorders();
                TopBorder topBorder286 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder414 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder298 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder414 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders412.Append(topBorder286);
                tableCellBorders412.Append(leftBorder414);
                tableCellBorders412.Append(bottomBorder298);
                tableCellBorders412.Append(rightBorder414);
                Shading shading412 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties412.Append(tableCellWidth412);
                tableCellProperties412.Append(verticalMerge38);
                tableCellProperties412.Append(tableCellBorders412);
                tableCellProperties412.Append(shading412);

                Paragraph paragraph444 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties444 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties444 = new ParagraphMarkRunProperties();
                RunFonts runFonts729 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize418 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript418 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties444.Append(runFonts729);
                paragraphMarkRunProperties444.Append(fontSize418);
                paragraphMarkRunProperties444.Append(fontSizeComplexScript418);

                paragraphProperties444.Append(paragraphMarkRunProperties444);

                paragraph444.Append(paragraphProperties444);

                tableCell412.Append(tableCellProperties412);
                tableCell412.Append(paragraph444);

                TableCell tableCell413 = new TableCell();

                TableCellProperties tableCellProperties413 = new TableCellProperties();
                TableCellWidth tableCellWidth413 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge39 = new VerticalMerge();

                TableCellBorders tableCellBorders413 = new TableCellBorders();
                TopBorder topBorder287 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder415 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder299 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder415 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders413.Append(topBorder287);
                tableCellBorders413.Append(leftBorder415);
                tableCellBorders413.Append(bottomBorder299);
                tableCellBorders413.Append(rightBorder415);
                Shading shading413 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties413.Append(tableCellWidth413);
                tableCellProperties413.Append(verticalMerge39);
                tableCellProperties413.Append(tableCellBorders413);
                tableCellProperties413.Append(shading413);

                Paragraph paragraph445 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties445 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties445 = new ParagraphMarkRunProperties();
                RunFonts runFonts730 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize419 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript419 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties445.Append(runFonts730);
                paragraphMarkRunProperties445.Append(fontSize419);
                paragraphMarkRunProperties445.Append(fontSizeComplexScript419);

                paragraphProperties445.Append(paragraphMarkRunProperties445);

                paragraph445.Append(paragraphProperties445);

                tableCell413.Append(tableCellProperties413);
                tableCell413.Append(paragraph445);

                TableCell tableCell414 = new TableCell();

                TableCellProperties tableCellProperties414 = new TableCellProperties();
                TableCellWidth tableCellWidth414 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge40 = new VerticalMerge();

                TableCellBorders tableCellBorders414 = new TableCellBorders();
                LeftBorder leftBorder416 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder300 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder416 = new RightBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders414.Append(leftBorder416);
                tableCellBorders414.Append(bottomBorder300);
                tableCellBorders414.Append(rightBorder416);
                Shading shading414 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties414.Append(tableCellWidth414);
                tableCellProperties414.Append(verticalMerge40);
                tableCellProperties414.Append(tableCellBorders414);
                tableCellProperties414.Append(shading414);

                Paragraph paragraph446 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties446 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties446 = new ParagraphMarkRunProperties();
                RunFonts runFonts731 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize420 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript420 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties446.Append(runFonts731);
                paragraphMarkRunProperties446.Append(fontSize420);
                paragraphMarkRunProperties446.Append(fontSizeComplexScript420);

                paragraphProperties446.Append(paragraphMarkRunProperties446);

                paragraph446.Append(paragraphProperties446);

                tableCell414.Append(tableCellProperties414);
                tableCell414.Append(paragraph446);

                TableCell tableCell415 = new TableCell();

                TableCellProperties tableCellProperties415 = new TableCellProperties();
                TableCellWidth tableCellWidth415 = new TableCellWidth() { Width = "720", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge41 = new VerticalMerge();

                TableCellBorders tableCellBorders415 = new TableCellBorders();
                LeftBorder leftBorder417 = new LeftBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder301 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder417 = new RightBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders415.Append(leftBorder417);
                tableCellBorders415.Append(bottomBorder301);
                tableCellBorders415.Append(rightBorder417);
                Shading shading415 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties415.Append(tableCellWidth415);
                tableCellProperties415.Append(verticalMerge41);
                tableCellProperties415.Append(tableCellBorders415);
                tableCellProperties415.Append(shading415);

                Paragraph paragraph447 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties447 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties447 = new ParagraphMarkRunProperties();
                RunFonts runFonts732 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize421 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript421 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties447.Append(runFonts732);
                paragraphMarkRunProperties447.Append(fontSize421);
                paragraphMarkRunProperties447.Append(fontSizeComplexScript421);

                paragraphProperties447.Append(paragraphMarkRunProperties447);

                paragraph447.Append(paragraphProperties447);

                tableCell415.Append(tableCellProperties415);
                tableCell415.Append(paragraph447);

                TableCell tableCell416 = new TableCell();

                TableCellProperties tableCellProperties416 = new TableCellProperties();
                TableCellWidth tableCellWidth416 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge42 = new VerticalMerge();

                TableCellBorders tableCellBorders416 = new TableCellBorders();
                LeftBorder leftBorder418 = new LeftBorder() { Val = BorderValues.Dashed, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder302 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder418 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders416.Append(leftBorder418);
                tableCellBorders416.Append(bottomBorder302);
                tableCellBorders416.Append(rightBorder418);
                Shading shading416 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties416.Append(tableCellWidth416);
                tableCellProperties416.Append(verticalMerge42);
                tableCellProperties416.Append(tableCellBorders416);
                tableCellProperties416.Append(shading416);

                Paragraph paragraph448 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties448 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties448 = new ParagraphMarkRunProperties();
                RunFonts runFonts733 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize422 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript422 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties448.Append(runFonts733);
                paragraphMarkRunProperties448.Append(fontSize422);
                paragraphMarkRunProperties448.Append(fontSizeComplexScript422);

                paragraphProperties448.Append(paragraphMarkRunProperties448);

                paragraph448.Append(paragraphProperties448);

                tableCell416.Append(tableCellProperties416);
                tableCell416.Append(paragraph448);

                TableCell tableCell417 = new TableCell();

                TableCellProperties tableCellProperties417 = new TableCellProperties();
                TableCellWidth tableCellWidth417 = new TableCellWidth() { Width = "2700", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan22 = new GridSpan() { Val = 4 };
                VerticalMerge verticalMerge43 = new VerticalMerge();

                TableCellBorders tableCellBorders417 = new TableCellBorders();
                LeftBorder leftBorder419 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder303 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder419 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders417.Append(leftBorder419);
                tableCellBorders417.Append(bottomBorder303);
                tableCellBorders417.Append(rightBorder419);
                Shading shading417 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties417.Append(tableCellWidth417);
                tableCellProperties417.Append(gridSpan22);
                tableCellProperties417.Append(verticalMerge43);
                tableCellProperties417.Append(tableCellBorders417);
                tableCellProperties417.Append(shading417);

                Paragraph paragraph449 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties449 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties449 = new ParagraphMarkRunProperties();
                RunFonts runFonts734 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize423 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript423 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties449.Append(runFonts734);
                paragraphMarkRunProperties449.Append(fontSize423);
                paragraphMarkRunProperties449.Append(fontSizeComplexScript423);

                paragraphProperties449.Append(paragraphMarkRunProperties449);

                paragraph449.Append(paragraphProperties449);

                tableCell417.Append(tableCellProperties417);
                tableCell417.Append(paragraph449);

                tableRow.Append(tableRowProperties4);
                tableRow.Append(tableCell82);
                tableRow.Append(tableCell83);
                tableRow.Append(tableCell84);
                tableRow.Append(tableCell85);
                tableRow.Append(tableCell86);
                tableRow.Append(tableCell87);
                tableRow.Append(tableCell88);
                tableRow.Append(tableCell89);
                tableRow.Append(tableCell90);
                tableRow.Append(tableCell91);
                tableRow.Append(tableCell411);
                tableRow.Append(tableCell412);
                tableRow.Append(tableCell413);
                tableRow.Append(tableCell414);
                tableRow.Append(tableCell415);
                tableRow.Append(tableCell416);
                tableRow.Append(tableCell417);

            }
            else if (index == 18 || (SheetNumber == 2 && (index - 34) == 18))
            {
                TableCell tableCell428 = new TableCell();

                TableCellProperties tableCellProperties428 = new TableCellProperties();
                TableCellWidth tableCellWidth428 = new TableCellWidth() { Width = "3168", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan23 = new GridSpan() { Val = 3 };
                VerticalMerge verticalMerge44 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders428 = new TableCellBorders();
                LeftBorder leftBorder430 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                RightBorder rightBorder430 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders428.Append(leftBorder430);
                tableCellBorders428.Append(rightBorder430);
                Shading shading428 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties428.Append(tableCellWidth428);
                tableCellProperties428.Append(gridSpan23);
                tableCellProperties428.Append(verticalMerge44);
                tableCellProperties428.Append(tableCellBorders428);
                tableCellProperties428.Append(shading428);

                Paragraph paragraph460 = new Paragraph() { RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties460 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties460 = new ParagraphMarkRunProperties();
                RunFonts runFonts770 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color22 = new Color() { Val = "808080" };
                FontSize fontSize458 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript458 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties460.Append(runFonts770);
                paragraphMarkRunProperties460.Append(color22);
                paragraphMarkRunProperties460.Append(fontSize458);
                paragraphMarkRunProperties460.Append(fontSizeComplexScript458);

                paragraphProperties460.Append(paragraphMarkRunProperties460);

                Run run312 = new Run() { RsidRunProperties = "00ED7FB2" };

                RunProperties runProperties312 = new RunProperties();
                RunFonts runFonts771 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color23 = new Color() { Val = "808080" };
                FontSize fontSize459 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript459 = new FontSizeComplexScript() { Val = "16" };

                runProperties312.Append(runFonts771);
                runProperties312.Append(color23);
                runProperties312.Append(fontSize459);
                runProperties312.Append(fontSizeComplexScript459);
                Text text312 = new Text();
                text312.Text = "R- Log’s";

                run312.Append(runProperties312);
                run312.Append(text312);

                paragraph460.Append(paragraphProperties460);
                paragraph460.Append(run312);

                Paragraph paragraph461 = new Paragraph() { RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties461 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties461 = new ParagraphMarkRunProperties();
                RunFonts runFonts772 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color24 = new Color() { Val = "808080" };
                FontSize fontSize460 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript460 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties461.Append(runFonts772);
                paragraphMarkRunProperties461.Append(color24);
                paragraphMarkRunProperties461.Append(fontSize460);
                paragraphMarkRunProperties461.Append(fontSizeComplexScript460);

                paragraphProperties461.Append(paragraphMarkRunProperties461);

                paragraph461.Append(paragraphProperties461);

                Paragraph paragraph462 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties462 = new ParagraphProperties();
                Justification justification115 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties462 = new ParagraphMarkRunProperties();
                RunFonts runFonts773 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize461 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript461 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties462.Append(runFonts773);
                paragraphMarkRunProperties462.Append(fontSize461);
                paragraphMarkRunProperties462.Append(fontSizeComplexScript461);

                paragraphProperties462.Append(justification115);
                paragraphProperties462.Append(paragraphMarkRunProperties462);

                Run run313 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties313 = new RunProperties();
                RunFonts runFonts774 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize462 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript462 = new FontSizeComplexScript() { Val = "20" };

                runProperties313.Append(runFonts774);
                runProperties313.Append(fontSize462);
                runProperties313.Append(fontSizeComplexScript462);
                Text text313 = new Text();
                text313.Text = labSheetA1Sheet.DuplicateRLog;

                run313.Append(runProperties313);
                run313.Append(text313);

                paragraph462.Append(paragraphProperties462);
                paragraph462.Append(run313);

                tableCell428.Append(tableCellProperties428);
                tableCell428.Append(paragraph460);
                tableCell428.Append(paragraph461);
                tableCell428.Append(paragraph462);

                TableCell tableCell429 = new TableCell();

                TableCellProperties tableCellProperties429 = new TableCellProperties();
                TableCellWidth tableCellWidth429 = new TableCellWidth() { Width = "2928", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan24 = new GridSpan() { Val = 5 };
                VerticalMerge verticalMerge45 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders429 = new TableCellBorders();
                LeftBorder leftBorder431 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder431 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders429.Append(leftBorder431);
                tableCellBorders429.Append(rightBorder431);
                Shading shading429 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties429.Append(tableCellWidth429);
                tableCellProperties429.Append(gridSpan24);
                tableCellProperties429.Append(verticalMerge45);
                tableCellProperties429.Append(tableCellBorders429);
                tableCellProperties429.Append(shading429);

                Paragraph paragraph463 = new Paragraph() { RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties463 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties463 = new ParagraphMarkRunProperties();
                RunFonts runFonts777 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color25 = new Color() { Val = "808080" };
                FontSize fontSize465 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript465 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties463.Append(runFonts777);
                paragraphMarkRunProperties463.Append(color25);
                paragraphMarkRunProperties463.Append(fontSize465);
                paragraphMarkRunProperties463.Append(fontSizeComplexScript465);

                paragraphProperties463.Append(paragraphMarkRunProperties463);

                Run run316 = new Run() { RsidRunProperties = "00ED7FB2" };

                RunProperties runProperties316 = new RunProperties();
                RunFonts runFonts778 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color26 = new Color() { Val = "808080" };
                FontSize fontSize466 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript466 = new FontSizeComplexScript() { Val = "16" };

                runProperties316.Append(runFonts778);
                runProperties316.Append(color26);
                runProperties316.Append(fontSize466);
                runProperties316.Append(fontSizeComplexScript466);
                Text text316 = new Text();
                text316.Text = "Precision Criteria";

                run316.Append(runProperties316);
                run316.Append(text316);

                paragraph463.Append(paragraphProperties463);
                paragraph463.Append(run316);

                Paragraph paragraph464 = new Paragraph() { RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties464 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties464 = new ParagraphMarkRunProperties();
                RunFonts runFonts779 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color27 = new Color() { Val = "808080" };
                FontSize fontSize467 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript467 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties464.Append(runFonts779);
                paragraphMarkRunProperties464.Append(color27);
                paragraphMarkRunProperties464.Append(fontSize467);
                paragraphMarkRunProperties464.Append(fontSizeComplexScript467);

                paragraphProperties464.Append(paragraphMarkRunProperties464);

                paragraph464.Append(paragraphProperties464);

                Paragraph paragraph465 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties465 = new ParagraphProperties();
                Justification justification116 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties465 = new ParagraphMarkRunProperties();
                RunFonts runFonts780 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color28 = new Color() { Val = "808080" };
                FontSize fontSize468 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript468 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties465.Append(runFonts780);
                paragraphMarkRunProperties465.Append(color28);
                paragraphMarkRunProperties465.Append(fontSize468);
                paragraphMarkRunProperties465.Append(fontSizeComplexScript468);

                paragraphProperties465.Append(justification116);
                paragraphProperties465.Append(paragraphMarkRunProperties465);

                Run run317 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties317 = new RunProperties();
                RunFonts runFonts781 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize469 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript469 = new FontSizeComplexScript() { Val = "20" };

                runProperties317.Append(runFonts781);
                runProperties317.Append(fontSize469);
                runProperties317.Append(fontSizeComplexScript469);
                Text text317 = new Text();
                text317.Text = labSheetA1Sheet.DuplicatePrecisionCriteria;

                run317.Append(runProperties317);
                run317.Append(text317);

                paragraph465.Append(paragraphProperties465);
                paragraph465.Append(run317);

                tableCell429.Append(tableCellProperties429);
                tableCell429.Append(paragraph463);
                tableCell429.Append(paragraph464);
                tableCell429.Append(paragraph465);

                TableCell tableCell430 = new TableCell();

                TableCellProperties tableCellProperties430 = new TableCellProperties();
                TableCellWidth tableCellWidth430 = new TableCellWidth() { Width = "1572", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan25 = new GridSpan() { Val = 2 };
                VerticalMerge verticalMerge46 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders430 = new TableCellBorders();
                LeftBorder leftBorder432 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder432 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders430.Append(leftBorder432);
                tableCellBorders430.Append(rightBorder432);
                Shading shading430 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties430.Append(tableCellWidth430);
                tableCellProperties430.Append(gridSpan25);
                tableCellProperties430.Append(verticalMerge46);
                tableCellProperties430.Append(tableCellBorders430);
                tableCellProperties430.Append(shading430);

                Paragraph paragraph466 = new Paragraph() { RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties466 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties466 = new ParagraphMarkRunProperties();
                RunFonts runFonts784 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color29 = new Color() { Val = "808080" };
                FontSize fontSize472 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript472 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties466.Append(runFonts784);
                paragraphMarkRunProperties466.Append(color29);
                paragraphMarkRunProperties466.Append(fontSize472);
                paragraphMarkRunProperties466.Append(fontSizeComplexScript472);

                paragraphProperties466.Append(paragraphMarkRunProperties466);

                Run run320 = new Run() { RsidRunProperties = "00ED7FB2" };

                RunProperties runProperties320 = new RunProperties();
                RunFonts runFonts785 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color30 = new Color() { Val = "808080" };
                FontSize fontSize473 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript473 = new FontSizeComplexScript() { Val = "16" };

                runProperties320.Append(runFonts785);
                runProperties320.Append(color30);
                runProperties320.Append(fontSize473);
                runProperties320.Append(fontSizeComplexScript473);
                Text text320 = new Text();
                text320.Text = "Acceptable";

                run320.Append(runProperties320);
                run320.Append(text320);

                paragraph466.Append(paragraphProperties466);
                paragraph466.Append(run320);

                Paragraph paragraph467 = new Paragraph() { RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties467 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties467 = new ParagraphMarkRunProperties();
                RunFonts runFonts786 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color31 = new Color() { Val = "808080" };
                FontSize fontSize474 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript474 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties467.Append(runFonts786);
                paragraphMarkRunProperties467.Append(color31);
                paragraphMarkRunProperties467.Append(fontSize474);
                paragraphMarkRunProperties467.Append(fontSizeComplexScript474);

                paragraphProperties467.Append(paragraphMarkRunProperties467);

                paragraph467.Append(paragraphProperties467);

                Paragraph paragraph468 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00027B3B", RsidParagraphProperties = "00027B3B", RsidRunAdditionDefault = "00027B3B" };

                ParagraphProperties paragraphProperties468 = new ParagraphProperties();
                Justification justification117 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties468 = new ParagraphMarkRunProperties();
                RunFonts runFonts787 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color32 = new Color() { Val = "808080" };
                FontSize fontSize475 = new FontSize() { Val = "16" };
                FontSizeComplexScript fontSizeComplexScript475 = new FontSizeComplexScript() { Val = "16" };

                paragraphMarkRunProperties468.Append(runFonts787);
                paragraphMarkRunProperties468.Append(color32);
                paragraphMarkRunProperties468.Append(fontSize475);
                paragraphMarkRunProperties468.Append(fontSizeComplexScript475);

                paragraphProperties468.Append(justification117);
                paragraphProperties468.Append(paragraphMarkRunProperties468);

                Run run321 = new Run() { RsidRunProperties = "0072154A" };

                RunProperties runProperties321 = new RunProperties();
                RunFonts runFonts788 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                FontSize fontSize476 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript476 = new FontSizeComplexScript() { Val = "20" };

                runProperties321.Append(runFonts788);
                runProperties321.Append(fontSize476);
                runProperties321.Append(fontSizeComplexScript476);
                Text text321 = new Text();
                text321.Text = labSheetA1Sheet.DuplicateAcceptableOrUnacceptable;

                run321.Append(runProperties321);
                run321.Append(text321);

                paragraph468.Append(paragraphProperties468);
                paragraph468.Append(run321);

                tableCell430.Append(tableCellProperties430);
                tableCell430.Append(paragraph466);
                tableCell430.Append(paragraph467);
                tableCell430.Append(paragraph468);

                tableRow.Append(tableRowProperties4);
                tableRow.Append(tableCell82);
                tableRow.Append(tableCell83);
                tableRow.Append(tableCell84);
                tableRow.Append(tableCell85);
                tableRow.Append(tableCell86);
                tableRow.Append(tableCell87);
                tableRow.Append(tableCell88);
                tableRow.Append(tableCell89);
                tableRow.Append(tableCell90);
                tableRow.Append(tableCell91);
                tableRow.Append(tableCell428);
                tableRow.Append(tableCell429);
                tableRow.Append(tableCell430);

            }
            else if (index == 19 || (SheetNumber == 2 && (index - 34) == 19))
            {
                TableCell tableCell441 = new TableCell();

                TableCellProperties tableCellProperties441 = new TableCellProperties();
                TableCellWidth tableCellWidth441 = new TableCellWidth() { Width = "3168", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan26 = new GridSpan() { Val = 3 };
                VerticalMerge verticalMerge47 = new VerticalMerge();

                TableCellBorders tableCellBorders441 = new TableCellBorders();
                LeftBorder leftBorder443 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder320 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder443 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders441.Append(leftBorder443);
                tableCellBorders441.Append(bottomBorder320);
                tableCellBorders441.Append(rightBorder443);
                Shading shading441 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties441.Append(tableCellWidth441);
                tableCellProperties441.Append(gridSpan26);
                tableCellProperties441.Append(verticalMerge47);
                tableCellProperties441.Append(tableCellBorders441);
                tableCellProperties441.Append(shading441);

                Paragraph paragraph479 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00801A19", RsidParagraphProperties = "00801A19", RsidRunAdditionDefault = "00801A19" };

                ParagraphProperties paragraphProperties479 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties479 = new ParagraphMarkRunProperties();
                RunFonts runFonts827 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                paragraphMarkRunProperties479.Append(runFonts827);

                paragraphProperties479.Append(paragraphMarkRunProperties479);

                paragraph479.Append(paragraphProperties479);

                tableCell441.Append(tableCellProperties441);
                tableCell441.Append(paragraph479);

                TableCell tableCell442 = new TableCell();

                TableCellProperties tableCellProperties442 = new TableCellProperties();
                TableCellWidth tableCellWidth442 = new TableCellWidth() { Width = "2928", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan27 = new GridSpan() { Val = 5 };
                VerticalMerge verticalMerge48 = new VerticalMerge();

                TableCellBorders tableCellBorders442 = new TableCellBorders();
                LeftBorder leftBorder444 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder321 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder444 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders442.Append(leftBorder444);
                tableCellBorders442.Append(bottomBorder321);
                tableCellBorders442.Append(rightBorder444);
                Shading shading442 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties442.Append(tableCellWidth442);
                tableCellProperties442.Append(gridSpan27);
                tableCellProperties442.Append(verticalMerge48);
                tableCellProperties442.Append(tableCellBorders442);
                tableCellProperties442.Append(shading442);

                Paragraph paragraph480 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00801A19", RsidParagraphProperties = "00801A19", RsidRunAdditionDefault = "00801A19" };

                ParagraphProperties paragraphProperties480 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties480 = new ParagraphMarkRunProperties();
                RunFonts runFonts828 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color33 = new Color() { Val = "808080" };
                FontSize fontSize514 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript514 = new FontSizeComplexScript() { Val = "18" };

                paragraphMarkRunProperties480.Append(runFonts828);
                paragraphMarkRunProperties480.Append(color33);
                paragraphMarkRunProperties480.Append(fontSize514);
                paragraphMarkRunProperties480.Append(fontSizeComplexScript514);

                paragraphProperties480.Append(paragraphMarkRunProperties480);

                paragraph480.Append(paragraphProperties480);

                tableCell442.Append(tableCellProperties442);
                tableCell442.Append(paragraph480);

                TableCell tableCell443 = new TableCell();

                TableCellProperties tableCellProperties443 = new TableCellProperties();
                TableCellWidth tableCellWidth443 = new TableCellWidth() { Width = "1572", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan28 = new GridSpan() { Val = 2 };
                VerticalMerge verticalMerge49 = new VerticalMerge();

                TableCellBorders tableCellBorders443 = new TableCellBorders();
                LeftBorder leftBorder445 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder322 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder445 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders443.Append(leftBorder445);
                tableCellBorders443.Append(bottomBorder322);
                tableCellBorders443.Append(rightBorder445);
                Shading shading443 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" };

                tableCellProperties443.Append(tableCellWidth443);
                tableCellProperties443.Append(gridSpan28);
                tableCellProperties443.Append(verticalMerge49);
                tableCellProperties443.Append(tableCellBorders443);
                tableCellProperties443.Append(shading443);

                Paragraph paragraph481 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00801A19", RsidParagraphProperties = "00801A19", RsidRunAdditionDefault = "00801A19" };

                ParagraphProperties paragraphProperties481 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties481 = new ParagraphMarkRunProperties();
                RunFonts runFonts829 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Color color34 = new Color() { Val = "808080" };
                FontSize fontSize515 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript515 = new FontSizeComplexScript() { Val = "18" };

                paragraphMarkRunProperties481.Append(runFonts829);
                paragraphMarkRunProperties481.Append(color34);
                paragraphMarkRunProperties481.Append(fontSize515);
                paragraphMarkRunProperties481.Append(fontSizeComplexScript515);

                paragraphProperties481.Append(paragraphMarkRunProperties481);

                paragraph481.Append(paragraphProperties481);

                tableCell443.Append(tableCellProperties443);
                tableCell443.Append(paragraph481);

                tableRow.Append(tableRowProperties4);
                tableRow.Append(tableCell82);
                tableRow.Append(tableCell83);
                tableRow.Append(tableCell84);
                tableRow.Append(tableCell85);
                tableRow.Append(tableCell86);
                tableRow.Append(tableCell87);
                tableRow.Append(tableCell88);
                tableRow.Append(tableCell89);
                tableRow.Append(tableCell90);
                tableRow.Append(tableCell91);
                tableRow.Append(tableCell441);
                tableRow.Append(tableCell442);
                tableRow.Append(tableCell443);

            }
            else
            {
            }

        }
    }
}