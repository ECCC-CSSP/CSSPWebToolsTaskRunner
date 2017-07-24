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
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public class DocxServiceMunicipality
    {
        #region Variables
        #endregion Variables

        #region Properties
        public GeneratedClassMunicipalityDocx _Docx { get; set; }
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public DocxServiceMunicipality(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _Docx = new GeneratedClassMunicipalityDocx(_TaskRunnerBaseService);
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

    public class GeneratedClassMunicipalityDocx
    {
        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public DocxServiceBase DocxBase { get; private set; }
        #endregion Properties

        #region Constructors
        public GeneratedClassMunicipalityDocx(TaskRunnerBaseService taskRunnerBaseService)
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
            DocxBase.CurrentTitle = DocxServiceMunicipalityRes.Title;

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

            GenerateMunicipalityDocument(document);

            mainDocumentPart1.Document = document;
        }
        private void GenerateMunicipalityDocument(Document document)
        {
            Body body = new Body();
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Table table = new Table();
            TableRow tableRow = new TableRow();
            TableCell tableCell = new TableCell();
            string URL = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemStatService tvItemStatService = new TVItemStatService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            TVItemModel tvItemModelMunicipality = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMunicipality.Error))
            {
                paragraph = DocxBase.AddParagraph(body);
                DocxBase.AddRunWithCurrentParagraphStyle(paragraph, tvItemModelMunicipality.Error);
            }

            //tvItemStatService.SetTVItemStatForTVItemIDAndParentsTVItemID(tvItemModelMunicipality.TVItemID);

            DocxBase.CurrentParagraphStyle = ParagraphStyleEnum.Heading1;
            DocxBase.CurrentJustificationValue = JustificationValues.Center;
            paragraph = DocxBase.AddParagraph(body);

            URL = _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelMunicipality);
            run = DocxBase.AddRunHyperlink(paragraph, URL, tvItemModelMunicipality.TVText);

            List<TVItemModel> tvItemModelInfrastructureList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Infrastructure);

            foreach (TVItemModel tvItemModelInfrastructure in tvItemModelInfrastructureList)
            {
                DocxBase.CurrentParagraphStyle = ParagraphStyleEnum.Heading2;
                paragraph = DocxBase.AddParagraph(body);

                URL = _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelInfrastructure);
                run = DocxBase.AddRunHyperlink(paragraph, URL, tvItemModelInfrastructure.TVText);

               
            }

            DocxBase.CurrentParagraphStyle = ParagraphStyleEnum.Normal;
            DocxBase.CurrentJustificationValue = JustificationValues.Left;
            paragraph = DocxBase.AddParagraph(body);

            paragraph = DocxBase.AddParagraph(body);

            DocxBase.AddRunWithCurrentParagraphStyle(paragraph, "Etc ... ");

            DocxBase.AddSectionProp(body);

            document.Append(body);

        }
    }
}