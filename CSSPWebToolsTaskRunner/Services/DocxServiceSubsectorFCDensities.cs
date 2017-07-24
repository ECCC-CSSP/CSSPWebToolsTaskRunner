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
using CSSPEnumsDLL.Services;

namespace CSSPWebToolsTaskRunner.Services
{
    public class DocxServiceSubsectorFaecalColiformDensities
    {
        #region Variables
        #endregion Variables

        #region Properties
        public GeneratedClassSubsectorFaecalColiformDensitiesDocx _Docx { get; set; }
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public DocxServiceSubsectorFaecalColiformDensities(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _Docx = new GeneratedClassSubsectorFaecalColiformDensitiesDocx(_TaskRunnerBaseService);
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

    public class GeneratedClassSubsectorFaecalColiformDensitiesDocx
    {
        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public DocxServiceBase DocxBase { get; private set; }
        #endregion Properties

        #region Constructors
        public GeneratedClassSubsectorFaecalColiformDensitiesDocx(TaskRunnerBaseService taskRunnerBaseService)
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

            GenerateSubsectorFaecalColiformDensitiesDocument(document);

            mainDocumentPart1.Document = document;
        }
        private void GenerateSubsectorFaecalColiformDensitiesDocument(Document document)
        {
            Body body = new Body();
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Table table = new Table();
            TableRow tableRow = new TableRow();
            TableCell tableCell = new TableCell();
            //string URL = "";

            BaseEnumService baseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemStatService tvItemStatService = new TVItemStatService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMSampleService mwqmSampleService = new MWQMSampleService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMSiteService mwqmSiteService = new MWQMSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TideSiteService tideSiteService = new TideSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TideDataValueService tideDataValueService = new TideDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                paragraph = DocxBase.AddParagraph(body);
                DocxBase.AddRunWithCurrentParagraphStyle(paragraph, tvItemModelSubsector.Error);
            }

            //tvItemStatService.SetTVItemStatForTVItemIDAndParentsTVItemID(tvItemModelSubsector.TVItemID);

            TVItemModel tvItemModelTideSite = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.TideSite).FirstOrDefault();
            if (tvItemModelTideSite == null)
                if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
                {
                    paragraph = DocxBase.AddParagraph(body);
                    DocxBase.AddRunWithCurrentParagraphStyle(paragraph, DocxServiceSubsectorFCDensitiesRes.CoundNotFindTideSite);
                    return;
                }

            List<TVItemModel> tvItemModelMWQMList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);

            List<StationDateValue> stationDateValueList = new List<StationDateValue>();
            List<DateTime> DateWithDataList = new List<DateTime>();
            List<string> StationList = new List<string>();

            foreach (TVItemModel tvItemModelMWQM in tvItemModelMWQMList)
            {
                string Station = mwqmSiteService.GetMWQMSiteModelWithMWQMSiteTVItemIDDB(tvItemModelMWQM.TVItemID).MWQMSiteNumber;

                StationList.Add(Station);

                List<MWQMSampleModel> mwqmSampleModelList = mwqmSampleService.GetMWQMSampleModelListWithMWQMSiteTVItemIDDB(tvItemModelMWQM.TVItemID);

                foreach (MWQMSampleModel mwqmSampleModel in mwqmSampleModelList.OrderByDescending(c => c.SampleDateTime_Local))
                {
                    StationDateValue stationDateValue = new StationDateValue()
                    {
                        Station = Station,
                        Date = new DateTime(mwqmSampleModel.SampleDateTime_Local.Year, mwqmSampleModel.SampleDateTime_Local.Month, mwqmSampleModel.SampleDateTime_Local.Day),
                        Value = mwqmSampleModel.FecCol_MPN_100ml,
                    };

                    if (!DateWithDataList.Contains(stationDateValue.Date))
                    {
                        DateWithDataList.Add(stationDateValue.Date);
                    }

                    stationDateValueList.Add(stationDateValue);
                }
            }

            DateWithDataList = DateWithDataList.OrderBy(c => c).ToList();
            StationList = StationList.OrderBy(c => c).ToList();

            for (int i = 0, count = DateWithDataList.Count; i < count; i = i + 15)
            {
                DocxBase.CurrentFontName = FontNameEnum.Arial;
                DocxBase.CurrentFontSize = 16;
                DocxBase.CurrentParagraphStyle = ParagraphStyleEnum.Caption;
                DocxBase.CurrentJustificationValue = JustificationValues.Left;
                paragraph = DocxBase.AddParagraph(body);
                paragraph = DocxBase.AddParagraph(body);

                string TableTitle = DocxServiceSubsectorFCDensitiesRes.Table + "1B-" + ((int)(i / 15) + 1).ToString() + "." + DocxServiceSubsectorFCDensitiesRes.FaecalColiformDensitiesMPNPer100 +
                    " " + DocxServiceSubsectorFCDensitiesRes.For + " " + tvItemModelSubsector.TVText;
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, TableTitle);

                DocxBase.CurrentTableStyle = TableStyleEnum.PlainTable1;
                DocxBase.CurrentFontName = FontNameEnum.Arial;
                DocxBase.CurrentFontSize = 10;
                table = DocxBase.AddTableStyle(body);
                tableRow = DocxBase.AddTableRow(table);

                List<DateTime> dateTimeNext15 = (from c in DateWithDataList
                                                 where c >= DateWithDataList[i]
                                                 select c).Take(15).ToList<DateTime>();

                tableCell = DocxBase.AddTableCell(tableRow);
                paragraph = DocxBase.AddTableCellParagraph(tableCell);
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, " ");
                paragraph = DocxBase.AddTableCellParagraph(tableCell);
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, DocxServiceSubsectorFCDensitiesRes.Station);

                foreach (DateTime dateTime in dateTimeNext15)
                {
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, dateTime.ToString("yyyy"));
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, dateTime.ToString("dd MMM"));
                }

                List<StationDateValue> StationDateValueDataList = (from c in stationDateValueList
                                                                   orderby c.Station
                                                                   select c).ToList<StationDateValue>();

                foreach (string station in StationList)
                {
                    tableRow = DocxBase.AddTableRow(table);

                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, station);

                    foreach (DateTime dateTime in dateTimeNext15)
                    {
                        StationDateValue stationDateValueForDate = (from c in stationDateValueList
                                                                    where c.Station == station
                                                                    && c.Date.Year == dateTime.Year
                                                                    && c.Date.Month == dateTime.Month
                                                                    && c.Date.Day == dateTime.Day
                                                                    select c).FirstOrDefault();

                        tableCell = DocxBase.AddTableCell(tableRow);
                        paragraph = DocxBase.AddTableCellParagraph(tableCell);
                        run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, (stationDateValueForDate == null ? " " : (stationDateValueForDate.Value < 2 ? "<2" : stationDateValueForDate.Value.ToString("F0"))));
                    }
                }

                // Tide Row
                tableRow = DocxBase.AddTableRow(table);

                tableCell = DocxBase.AddTableCell(tableRow);
                paragraph = DocxBase.AddTableCellParagraph(tableCell);
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, DocxServiceSubsectorFCDensitiesRes.Tide);

                foreach (DateTime dateTime in dateTimeNext15)
                {
                    TideDataValueModel tideDataValueModel = tideDataValueService.GetTideDataValueModelWithTideSiteTVItemIDAndDateDB(tvItemModelTideSite.TVItemID, dateTime);

                    string TideStartAccronym = GetTideTextAccronym(tideDataValueModel.TideStart);
                    string TideEndAccronym = GetTideTextAccronym(tideDataValueModel.TideEnd);

                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, TideStartAccronym + "-" + TideEndAccronym);
                }

                // Rain (mm) Row
                tableRow = DocxBase.AddTableRow(table);

                tableCell = DocxBase.AddTableCell(tableRow);
                paragraph = DocxBase.AddTableCellParagraph(tableCell);
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, DocxServiceSubsectorFCDensitiesRes.Rain_mm);

                foreach (DateTime dateTime in dateTimeNext15)
                {
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, "");
                }

                // 0-24 h Row
                tableRow = DocxBase.AddTableRow(table);

                tableCell = DocxBase.AddTableCell(tableRow);
                paragraph = DocxBase.AddTableCellParagraph(tableCell);
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, DocxServiceSubsectorFCDensitiesRes._0_24h);

                foreach (DateTime dateTime in dateTimeNext15)
                {
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, "ToDo");
                }

                // 0-48 h Row
                tableRow = DocxBase.AddTableRow(table);

                tableCell = DocxBase.AddTableCell(tableRow);
                paragraph = DocxBase.AddTableCellParagraph(tableCell);
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, DocxServiceSubsectorFCDensitiesRes._0_48h);

                foreach (DateTime dateTime in dateTimeNext15)
                {
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, "ToDo");
                }

                // 0-72 h Row
                tableRow = DocxBase.AddTableRow(table);

                tableCell = DocxBase.AddTableCell(tableRow);
                paragraph = DocxBase.AddTableCellParagraph(tableCell);
                run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, DocxServiceSubsectorFCDensitiesRes._0_72h);

                foreach (DateTime dateTime in dateTimeNext15)
                {
                    tableCell = DocxBase.AddTableCell(tableRow);
                    paragraph = DocxBase.AddTableCellParagraph(tableCell);
                    run = DocxBase.AddRunWithCurrentParagraphStyle(paragraph, "ToDo");
                }

            }


            DocxBase.CurrentParagraphStyle = ParagraphStyleEnum.Quote;
            paragraph = DocxBase.AddParagraph(body);
            string NoteTxt = DocxServiceSubsectorFCDensitiesRes.Note + " : ";
            for (int i = 1, count = Enum.GetNames(typeof(TideTextEnum)).Length; i < count; i++)
            {
                NoteTxt = GetTideTextAccronym((TideTextEnum)i) + " = " + baseEnumService.GetEnumText_TideTextEnum((TideTextEnum)i) + " | ";
            }

            DocxBase.AddRunWithCurrentParagraphStyle(paragraph, NoteTxt);

            DocxBase.AddSectionProp(body);

            document.Append(body);

        }

        private string GetTideTextAccronym(TideTextEnum? TideText)
        {
            if (TideText == null)
                return "";

            switch (TideText)
            {
                case TideTextEnum.Error:
                    return "";
                case TideTextEnum.HighTide:
                    return "HT";
                case TideTextEnum.HighTideFalling:
                    return "HF";
                case TideTextEnum.HighTideRising:
                    return "HR";
                case TideTextEnum.LowTide:
                    return "LT";
                case TideTextEnum.LowTideFalling:
                    return "LF";
                case TideTextEnum.LowTideRising:
                    return "LR";
                case TideTextEnum.MidTide:
                    return "MT";
                case TideTextEnum.MidTideFalling:
                    return "MF";
                case TideTextEnum.MidTideRising:
                    return "MR";
                default:
                    return "";
            }
        }
    }

    public class StationDateValue
    {
        public StationDateValue()
        {
        }
        public string Station { get; set; }
        public DateTime Date { get; set; }
        public int Value { get; set; }
    }
}