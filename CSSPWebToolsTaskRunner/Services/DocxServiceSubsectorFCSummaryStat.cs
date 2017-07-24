using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Linq;
using System;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;

namespace CSSPWebToolsTaskRunner.Services
{
    public class DocxServiceSubsectorFaecalColiformSummaryStat
    {
        #region Variables
        #endregion Variables

        #region Properties
        public GeneratedClassSubsectorFaecalColiformSummaryStatDocx _Docx { get; set; }
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public DocxServiceSubsectorFaecalColiformSummaryStat(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _Docx = new GeneratedClassSubsectorFaecalColiformSummaryStatDocx(_TaskRunnerBaseService);
        }
        #endregion Constructors

        public bool Generate(FileInfo fi)
        {
            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == "fr")
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
            }
            _Docx.CreatePackage(fi.FullName);

            return true;
        }
    }

    public class GeneratedClassSubsectorFaecalColiformSummaryStatDocx
    {
        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public DocxServiceBase DocxBase { get; private set; }
        #endregion Properties

        #region Constructors
        public GeneratedClassSubsectorFaecalColiformSummaryStatDocx(TaskRunnerBaseService taskRunnerBaseService)
        {
            DocxBase = new DocxServiceBase(taskRunnerBaseService);
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions
        #endregion Functions

        public void CreatePackage(string filePath)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        private void CreateParts(WordprocessingDocument document)
        {
            DocxBase.CurrentTitle = DocxServiceSubsectorRes.Title;

            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            DocxBase.GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId13");
            DocxBase.GenerateFontTablePart1Content(fontTablePart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId3");
            DocxBase.GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId7");
            DocxBase.GenerateEndnotesPart1Content(endnotesPart1);

            FooterPart footerPart1 = mainDocumentPart1.AddNewPart<FooterPart>("rId12");
            DocxBase.GenerateFooterPart1Content(footerPart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId2");
            DocxBase.GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            CustomXmlPart customXmlPart1 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId1");
            DocxBase.GenerateCustomXmlPart1Content(customXmlPart1);

            CustomXmlPropertiesPart customXmlPropertiesPart1 = customXmlPart1.AddNewPart<CustomXmlPropertiesPart>("rId1");
            DocxBase.GenerateCustomXmlPropertiesPart1Content(customXmlPropertiesPart1);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId6");
            DocxBase.GenerateFootnotesPart1Content(footnotesPart1);

            HeaderPart headerPart1 = mainDocumentPart1.AddNewPart<HeaderPart>("rId11");
            DocxBase.GenerateHeaderPart1Content(headerPart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId5");
            DocxBase.GenerateWebSettingsPart1Content(webSettingsPart1);

            //ChartPart chartPart1 = mainDocumentPart1.AddNewPart<ChartPart>("rId10");
            //DocxBase.GenerateChartPart1Content(chartPart1);

            //EmbeddedPackagePart embeddedPackagePart1 = chartPart1.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId3");
            //DocxBase.GenerateEmbeddedPackagePart1Content(embeddedPackagePart1);

            //ChartColorStylePart chartColorStylePart1 = chartPart1.AddNewPart<ChartColorStylePart>("rId2");
            //DocxBase.GenerateChartColorStylePart1Content(chartColorStylePart1);

            //ChartStylePart chartStylePart1 = chartPart1.AddNewPart<ChartStylePart>("rId1");
            //DocxBase.GenerateChartStylePart1Content(chartStylePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId4");
            DocxBase.GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            //ChartPart chartPart2 = mainDocumentPart1.AddNewPart<ChartPart>("rId9");
            //DocxBase.GenerateChartPart2Content(chartPart2);

            //EmbeddedPackagePart embeddedPackagePart2 = chartPart2.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId3");
            //DocxBase.GenerateEmbeddedPackagePart2Content(embeddedPackagePart2);

            //ChartColorStylePart chartColorStylePart2 = chartPart2.AddNewPart<ChartColorStylePart>("rId2");
            //DocxBase.GenerateChartColorStylePart2Content(chartColorStylePart2);

            //ChartStylePart chartStylePart2 = chartPart2.AddNewPart<ChartStylePart>("rId1");
            //DocxBase.GenerateChartStylePart2Content(chartStylePart2);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId14");
            DocxBase.GenerateThemePart1Content(themePart1);

            foreach (UsedHyperlink usedHyperlink in DocxBase.UsedHyperlinkList)
            {
                mainDocumentPart1.AddHyperlinkRelationship(new System.Uri(usedHyperlink.URL, System.UriKind.Absolute), true, usedHyperlink.Id.ToString());
            }

            DocxBase.SetPackageProperties(document);
        }
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            document.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            GenerateSubsectorFaecalColiformSummaryStatDocument(document);

            mainDocumentPart1.Document = document;
        }
        private void GenerateSubsectorFaecalColiformSummaryStatDocument(Document document)
        {
            Body body = new Body();
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Table table = new Table();
            TableRow tableRow = new TableRow();
            TableCell tableCell = new TableCell();
            //string URL = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemStatService tvItemStatService = new TVItemStatService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                paragraph = DocxBase.AddParagraph(body);
                DocxBase.AddRunWithCurrentParagraphStyle(paragraph, tvItemModelSubsector.Error);
            }

            //tvItemStatService.SetTVItemStatForTVItemIDAndParentsTVItemID(tvItemModelSubsector.TVItemID);

            DocxBase.CurrentParagraphStyle = ParagraphStyleEnum.Caption;
            DocxBase.CurrentJustificationValue = JustificationValues.Center;
            paragraph = DocxBase.AddParagraph(body);

            DocxBase.CurrentFontSize = 36;

            string TableTitle = DocxServiceSubsectorFCSummaryStatRes.Table + "1A-1." + DocxServiceSubsectorFCSummaryStatRes.SummaryStatOfFCDensitiesMPNPer100 + 
                " " + DocxServiceSubsectorFCSummaryStatRes.For + " " + tvItemModelSubsector.TVText;
            run = DocxBase.AddRunWithCurrentFontStyle(paragraph, TableTitle);

            DocxBase.CurrentTableStyle = TableStyleEnum.ListTable7Colorful_Accent5;
            table = DocxBase.AddTableStyle(body);
            tableRow = DocxBase.AddTableRow(table);

            List<string> ColumnTitleList = new List<string>()
            {
                DocxServiceSubsectorFCSummaryStatRes.Station,
                DocxServiceSubsectorFCSummaryStatRes.Samples,
                DocxServiceSubsectorFCSummaryStatRes.Period,
                DocxServiceSubsectorFCSummaryStatRes.MinFC,
                DocxServiceSubsectorFCSummaryStatRes.MaxFC,
                DocxServiceSubsectorFCSummaryStatRes.GMean,
                DocxServiceSubsectorFCSummaryStatRes.Median,
                DocxServiceSubsectorFCSummaryStatRes.P90,
                DocxServiceSubsectorFCSummaryStatRes.PercBigger43,
            };

            // Doing Cell Title
            foreach (string cellTitle in ColumnTitleList)
            {
                tableCell = DocxBase.AddTableCell(tableRow);
                paragraph = DocxBase.AddTableCellParagraph(tableCell);
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, cellTitle);
            }

            List<TVItemModel> tvItemModelMWQMList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);

            foreach (TVItemModel tvItemModelMWQM in tvItemModelMWQMList)
            {
                TVItemMoreInfoMWQMSiteModel tvItemMoreInfoMWQMSiteModel = new TVItemMoreInfoMWQMSiteModel();
                tvItemMoreInfoMWQMSiteModel = tvItemService.GetTVItemMoreInfoMWQMSiteTVItemIDDB(tvItemModelMWQM.TVItemID, 30);

                if (tvItemMoreInfoMWQMSiteModel.StatMaxYear > DateTime.Now.Year - 6)
                {
                    tableRow = DocxBase.AddTableRow(table);

                    // Doing Station
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, tvItemModelMWQM.TVText);

                    // Doing Samples
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    string SampCountTxt = (tvItemMoreInfoMWQMSiteModel.SampCount == null ? "" : tvItemMoreInfoMWQMSiteModel.SampCount.ToString());
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, SampCountTxt);

                    // Doing Period
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    string StatMinMaxYearTxt = tvItemMoreInfoMWQMSiteModel.StatMinYear.ToString() + "-" + tvItemMoreInfoMWQMSiteModel.StatMaxYear.ToString();
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, StatMinMaxYearTxt);

                    // Doing MinFC
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    string MinFCTxt = (tvItemMoreInfoMWQMSiteModel.MinFC == null ? "" : ((float)tvItemMoreInfoMWQMSiteModel.MinFC).ToString("F0"));
                    if (MinFCTxt == "1")
                        MinFCTxt = "< 2";
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, MinFCTxt);

                    // Doing MaxFC
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    string MaxFCTxt = (tvItemMoreInfoMWQMSiteModel.MaxFC == null ? "" : ((float)tvItemMoreInfoMWQMSiteModel.MaxFC).ToString("F0"));
                    if (MaxFCTxt == "1")
                        MaxFCTxt = "< 2";
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, MaxFCTxt);

                    // Doing GMean
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    string GMeanTxt = (tvItemMoreInfoMWQMSiteModel.GeoMean == null ? "" : ((float)tvItemMoreInfoMWQMSiteModel.GeoMean).ToString("F0"));
                    if (tvItemMoreInfoMWQMSiteModel.GeoMean > 14)
                    {
                        int? TempFS = (int)DocxBase.CurrentFontSize;
                        DocxBase.CurrentHighlightColorValue = HighlightColorValues.Yellow;
                        DocxBase.CurrentFontSize = 22;
                        run = DocxBase.AddRunWithCurrentFontStyle(paragraph, GMeanTxt);
                        DocxBase.CurrentHighlightColorValue = HighlightColorValues.None;
                        DocxBase.CurrentFontSize = TempFS;
                    }
                    else
                    {
                        run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, GMeanTxt);
                    }

                    // Doing Median
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    string MedianTxt = (tvItemMoreInfoMWQMSiteModel.Median == null ? "" : ((float)tvItemMoreInfoMWQMSiteModel.Median).ToString("F0"));
                    if (tvItemMoreInfoMWQMSiteModel.Median > 14)
                    {
                        int? TempFS = (int)DocxBase.CurrentFontSize;
                        DocxBase.CurrentHighlightColorValue = HighlightColorValues.Yellow;
                        DocxBase.CurrentFontSize = 22;
                        run = DocxBase.AddRunWithCurrentFontStyle(paragraph, MedianTxt);
                        DocxBase.CurrentHighlightColorValue = HighlightColorValues.None;
                        DocxBase.CurrentFontSize = TempFS;
                    }
                    else
                    {
                        run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, MedianTxt);
                    }

                    // Doing P90
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    string P90Txt = (tvItemMoreInfoMWQMSiteModel.P90 == null ? "" : ((float)tvItemMoreInfoMWQMSiteModel.P90).ToString("F0"));
                    if (tvItemMoreInfoMWQMSiteModel.P90 > 43)
                    {
                        int? TempFS = (int)DocxBase.CurrentFontSize;
                        DocxBase.CurrentHighlightColorValue = HighlightColorValues.Yellow;
                        DocxBase.CurrentFontSize = 22;
                        run = DocxBase.AddRunWithCurrentFontStyle(paragraph, P90Txt);
                        DocxBase.CurrentHighlightColorValue = HighlightColorValues.None;
                        DocxBase.CurrentFontSize = TempFS;
                    }
                    else
                    {
                        run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, P90Txt);
                    }

                    // Doing % > 43
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    string PercOver43Txt = (tvItemMoreInfoMWQMSiteModel.PercOver43 == null ? "" : ((float)tvItemMoreInfoMWQMSiteModel.PercOver43).ToString("F0"));
                    if (tvItemMoreInfoMWQMSiteModel.PercOver43 > 10)
                    {
                        int? TempFS = (int)DocxBase.CurrentFontSize;
                        DocxBase.CurrentHighlightColorValue = HighlightColorValues.Yellow;
                        DocxBase.CurrentFontSize = 22;
                        run = DocxBase.AddRunWithCurrentFontStyle(paragraph, PercOver43Txt);
                        DocxBase.CurrentHighlightColorValue = HighlightColorValues.None;
                        DocxBase.CurrentFontSize = TempFS;
                    }
                    else
                    {
                        run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, PercOver43Txt);
                    }
                }
            }

            DocxBase.CurrentParagraphStyle = ParagraphStyleEnum.Quote;
            paragraph = DocxBase.AddParagraph(body);
            string NoteTxt = DocxServiceSubsectorFCSummaryStatRes.Note + " : " + DocxServiceSubsectorFCSummaryStatRes.TheFollowingValuesHaveBeenShaded
                + " : " + DocxServiceSubsectorFCSummaryStatRes.GeometricMeanBigger14 + ", " + DocxServiceSubsectorFCSummaryStatRes.MedianBigger14 + ", "
                + DocxServiceSubsectorFCSummaryStatRes.PercBigger43More10Perc + ", " + DocxServiceSubsectorFCSummaryStatRes.Perc90Bigger43;

            DocxBase.AddRunWithCurrentParagraphStyle(paragraph, NoteTxt);

            DocxBase.AddSectionProp(body);

            document.Append(body);

        }
    }
}