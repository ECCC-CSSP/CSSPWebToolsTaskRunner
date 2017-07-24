using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Linq;
using CSSPWebToolsTaskRunner.Services.Resources;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Services;

namespace CSSPWebToolsTaskRunner.Services
{
    public class XlsxServiceSubsector
    {
        #region Variables
        #endregion Variables

        #region Properties
        public GeneratedClassSubsectorXlsx _Xlsx { get; set; }
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public XlsxServiceSubsector(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _Xlsx = new GeneratedClassSubsectorXlsx(_TaskRunnerBaseService);
        }
        #endregion Constructors

        public void GenerateSubsector(FileInfo fi)
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
            _Xlsx.CreatePackageSubsector(fi.FullName);
        }
        public void GenerateSubsectorPollutionSourceFieldSheet(FileInfo fi)
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
            _Xlsx.CreatePackageSubsectorPollutionSourceFieldSheet(fi.FullName);
        }
    }

    public class GeneratedClassSubsectorXlsx
    {
        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public XlsxServiceBase XlsxBase { get; private set; }
        #endregion Properties

        #region Constructors
        public GeneratedClassSubsectorXlsx(TaskRunnerBaseService taskRunnerBaseService)
        {
            XlsxBase = new XlsxServiceBase(taskRunnerBaseService);
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions
        #endregion Functions

        // Subsector
        public void CreatePackageSubsector(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsSubsector(package);
            }
        }
        private void CreatePartsSubsector(SpreadsheetDocument document)
        {
            int SheetOrdinal = 0;

            BaseEnumService baseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            XlsxBase.AddSheetNameAndID(baseEnumService.GetEnumText_TVTypeEnum(TVTypeEnum.Subsector));

            ExtendedFilePropertiesPart extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>("rId1000");
            XlsxBase.GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            XlsxBase.GenerateWorkbookPart1Content(workbookPart);

            SheetOrdinal = 0;
            WorksheetPart worksheetA = workbookPart.AddNewPart<WorksheetPart>(XlsxBase.sheetNameAndIDList[SheetOrdinal].SheetID);
            GenerateWorksheetAContentSubsector(worksheetA, SheetOrdinal);

            foreach (UsedHyperlink usedHyperlink in XlsxBase.UsedHyperlinkList)
            {
                worksheetA.AddHyperlinkRelationship(new System.Uri(usedHyperlink.URL, System.UriKind.Absolute), true, usedHyperlink.Id);
            }

            SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>("rId2000");
            XlsxBase.GenerateSharedStringTablePart1Content(sharedStringTablePart);

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3000");
            XlsxBase.GenerateWorkbookStylesPart1Content(workbookStylesPart);

            ThemePart themePart = workbookPart.AddNewPart<ThemePart>("rId4000");
            XlsxBase.GenerateThemePart1Content(themePart);


            XlsxBase.SetPackageProperties(document);

        }
        private void GenerateWorksheetAContentSubsector(WorksheetPart worksheetPart, int SheetOrdinal)
        {
            MergeCells mergeCells = new MergeCells();
            Row row = new Row();
            Cell cell = new Cell();
            Hyperlinks hyperlinks = new Hyperlinks();
            string Id = XlsxBase.sheetNameAndIDList[SheetOrdinal].SheetID;

            XlsxBase.CurrentColumn = 0;
            XlsxBase.CurrentRow = 0;
            XlsxBase.CurrentColumnProp = 0;
            Worksheet worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetViews sheetViews = new SheetViews();

            SheetView sheetView = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView.Append(selection);

            sheetViews.Append(sheetView);
            SheetFormatProperties sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns = new Columns();
            List<double?> columnWidthList = new List<double?>() { 30D };
            foreach (double? width in columnWidthList)
            {
                Column colum = XlsxBase.AddColumnProp(columns, width);
            }

            SheetData sheetData = new SheetData();

            BaseEnumService baseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemStatService tvItemStatService = new TVItemStatService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            List<TVItemModel> tvItemModelMunicipalityList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Municipality);

            foreach (TVItemModel tvItemModelMunicipality in tvItemModelMunicipalityList)
            {
                //tvItemStatService.SetTVItemStatForTVItemIDAndParentsTVItemID(tvItemModelMunicipality.TVItemID);

                XlsxBase.CurrentRowHeight = 24D;
                XlsxBase.CurrentFontSize = 18;
                XlsxBase.CurrentBorderStyleValue = BorderStyleValues.Thick;
                XlsxBase.CurrentBottomBorder = true;
                XlsxBase.CurrentBorderColor = System.Drawing.Color.Green;
                XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

                row = XlsxBase.AddRow();
                string URL = _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelMunicipality);
                XlsxBase.CurrentFontColor = System.Drawing.Color.Blue;
                cell = XlsxBase.AddCellHyperlink(hyperlinks, row, tvItemModelMunicipality.TVText, URL);
                XlsxBase.CurrentFontColor = null;
                //cell = XlsxBase.AddCellString(row, tvItemModel.TVText);
                cell = XlsxBase.AddCellString(row, null);
                sheetData.Append(row);

                MergeCell mergeCell = new MergeCell() { Reference = "A" + XlsxBase.CurrentRow.ToString() + ":B" + XlsxBase.CurrentRow.ToString() };

                mergeCells.Append(mergeCell);

                XlsxBase.CurrentRowHeight = 16D;
                XlsxBase.CurrentFontSize = 12;
                XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

                List<TVItemStatModel> tvItemStatModelList = tvItemStatService.GetTVItemStatModelListWithTVItemIDDB(tvItemModelMunicipality.TVItemID);

                int count = 0;
                foreach (TVItemStatModel tvItemStatModel in tvItemStatModelList)
                {
                    count += 1;
                    row = XlsxBase.AddRow();

                    if (count % 5 == 0)
                    {
                        XlsxBase.CurrentBorderStyleValue = BorderStyleValues.Thin;
                        XlsxBase.CurrentBottomBorder = true;
                    }
                    else
                    {
                        XlsxBase.CurrentBorderStyleValue = null;
                        XlsxBase.CurrentBottomBorder = false;
                    }
                    XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Right;
                    cell = XlsxBase.AddCellString(row, baseEnumService.GetEnumText_TVTypeEnum(tvItemStatModel.TVType));

                    XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;
                    cell = XlsxBase.AddCellNumber(row, tvItemStatModel.ChildCount.ToString());

                    sheetData.Append(row);
                }

                XlsxBase.CurrentBorderStyleValue = null;
                XlsxBase.CurrentBottomBorder = false;

                for (int i = 0; i < 2; i++)
                {
                    row = XlsxBase.AddRow();
                    cell = XlsxBase.AddCellString(row, null);
                    sheetData.Append(row);
                }
            }

            PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup = new PageSetup() { Orientation = OrientationValues.Portrait, Id = "rId" + SheetOrdinal.ToString() };

            worksheet.Append(sheetViews);
            worksheet.Append(sheetFormatProperties);

            if (columns.ChildElements.Count > 0)
                worksheet.Append(columns);

            worksheet.Append(sheetData);

            mergeCells.Count = (UInt32Value)((UInt32)mergeCells.ChildElements.Count);
            if (mergeCells.ChildElements.Count > 0)
                worksheet.Append(mergeCells);


            if (XlsxBase.UsedHyperlinkList.Count > 0)
                worksheet.Append(hyperlinks);

            worksheet.Append(pageMargins);
            worksheet.Append(pageSetup);

            worksheetPart.Worksheet = worksheet;

        }


        // SubsectorPollutionSourceFieldSheet
        public void CreatePackageSubsectorPollutionSourceFieldSheet(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsSubsectorPollutionSourceFieldSheet(package);
            }
        }
        private void CreatePartsSubsectorPollutionSourceFieldSheet(SpreadsheetDocument document)
        {
            int SheetOrdinal = 0;

            BaseEnumService baseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<string> SheetNameList = new List<string>();
            SheetNameList.Add("Active - " + baseEnumService.GetEnumText_TVTypeEnum(TVTypeEnum.Subsector));
            SheetNameList.Add("Inactive - " + baseEnumService.GetEnumText_TVTypeEnum(TVTypeEnum.Subsector));
            XlsxBase.AddSheetNameAndID(SheetNameList[0]);
            XlsxBase.AddSheetNameAndID(SheetNameList[1]);

            ExtendedFilePropertiesPart extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>("rId1000");
            XlsxBase.GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            XlsxBase.GenerateWorkbookPart1Content(workbookPart);

            SheetOrdinal = 0;
            WorksheetPart worksheetA = workbookPart.AddNewPart<WorksheetPart>(XlsxBase.sheetNameAndIDList[SheetOrdinal].SheetID);
            GenerateWorksheetAContentSubsectorPollutionSourceFieldSheet(worksheetA, workbookPart.Workbook, SheetNameList[0], SheetOrdinal, true);

            foreach (UsedHyperlink usedHyperlink in XlsxBase.UsedHyperlinkList)
            {
                worksheetA.AddHyperlinkRelationship(new System.Uri(usedHyperlink.URL, System.UriKind.Absolute), true, usedHyperlink.Id);
            }

            SheetOrdinal = 1;
            WorksheetPart worksheetB = workbookPart.AddNewPart<WorksheetPart>(XlsxBase.sheetNameAndIDList[SheetOrdinal].SheetID);
            GenerateWorksheetAContentSubsectorPollutionSourceFieldSheet(worksheetB, workbookPart.Workbook, SheetNameList[1], SheetOrdinal, false);

            foreach (UsedHyperlink usedHyperlink in XlsxBase.UsedHyperlinkList)
            {
                worksheetB.AddHyperlinkRelationship(new System.Uri(usedHyperlink.URL, System.UriKind.Absolute), true, usedHyperlink.Id);
            }

            SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>("rId2000");
            XlsxBase.GenerateSharedStringTablePart1Content(sharedStringTablePart);

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3000");
            XlsxBase.GenerateWorkbookStylesPart1Content(workbookStylesPart);

            ThemePart themePart = workbookPart.AddNewPart<ThemePart>("rId4000");
            XlsxBase.GenerateThemePart1Content(themePart);


            XlsxBase.SetPackageProperties(document);

        }
        private void GenerateWorksheetAContentSubsectorPollutionSourceFieldSheet(WorksheetPart worksheetPart, Workbook workbook, string SheetName, int SheetOrdinal, bool ActivePollutionSource)
        {
            BaseEnumService baseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            PolSourceSiteService polSourceSiteService = new PolSourceSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            //PolSourceObservationService polSourceObservationService = new PolSourceObservationService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            //MWQMSubsectorService mwqmSubsectorService = new MWQMSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            MergeCells mergeCells = new MergeCells();
            Row row = new Row();
            Cell cell = new Cell();
            Hyperlinks hyperlinks = new Hyperlinks();
            string Id = XlsxBase.sheetNameAndIDList[SheetOrdinal].SheetID;

            XlsxBase.CurrentColumn = 0;
            XlsxBase.CurrentRow = 0;
            XlsxBase.CurrentColumnProp = 0;
            Worksheet worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetViews sheetViews = new SheetViews();

            SheetView sheetView = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView.Append(selection);

            sheetViews.Append(sheetView);
            SheetFormatProperties sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns = new Columns();
            List<double?> columnWidthList = new List<double?>() { 7D, 10D, 90D, 12D, 12D, 10D, 14D, 20D };
            foreach (double? width in columnWidthList)
            {
                Column colum = XlsxBase.AddColumnProp(columns, width);
            }

            SheetData sheetData = new SheetData();

            TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            //MWQMSubsectorModel mwqmSubsectorModel = mwqmSubsectorService.GetMWQMSubsectorModelWithMWQMSubsectorTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            XlsxBase.CurrentRowHeight = 28D;
            XlsxBase.CurrentFontSize = 24;
            XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;
            XlsxBase.CurrentBorderStyleValue = BorderStyleValues.Thin;
            XlsxBase.CurrentBottomBorder = true;

            row = XlsxBase.AddRow();
            string URL = _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelSubsector);
            XlsxBase.CurrentFontColor = System.Drawing.Color.Blue;
            cell = XlsxBase.AddCellHyperlink(hyperlinks, row, tvItemModelSubsector.TVText, URL);
            XlsxBase.CurrentFontColor = null;
            //cell = XlsxBase.AddCellString(row, tvItemModel.TVText);
            cell = XlsxBase.AddCellString(row, null);
            sheetData.Append(row);

            MergeCell mergeCell = new MergeCell() { Reference = "A" + XlsxBase.CurrentRow.ToString() + ":H" + XlsxBase.CurrentRow.ToString() };
            mergeCells.Append(mergeCell);

            List<TVItemModel> tvItemModelPolSourceList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, TVTypeEnum.PolSourceSite);

            XlsxBase.CurrentRowHeight = 24D;
            XlsxBase.CurrentFontSize = 18;
            XlsxBase.CurrentBorderStyleValue = BorderStyleValues.Thin;
            XlsxBase.CurrentBottomBorder = true;
            XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

            row = XlsxBase.AddRow();
            XlsxBase.CurrentFontColor = null;
            cell = XlsxBase.AddCellString(row, "Site");
            cell = XlsxBase.AddCellString(row, "Type");
            cell = XlsxBase.AddCellString(row, "Observation");
            cell = XlsxBase.AddCellString(row, "Lat");
            cell = XlsxBase.AddCellString(row, "Lng");
            cell = XlsxBase.AddCellString(row, "Active");
            cell = XlsxBase.AddCellString(row, "Update");
            cell = XlsxBase.AddCellString(row, "Civic Address");
            sheetData.Append(row);

            XlsxBase.CurrentRowHeight = 160D;
            XlsxBase.CurrentFontSize = 12;
            XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Left;

            List<PolSourceSiteModel> polSourceSiteModelList = polSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);

            int countPolSourceSite = 0;
            foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList.OrderBy(c => c.Site).ToList())
            {
                bool IsActive = (from c in tvItemModelPolSourceList
                                 where c.TVItemID == polSourceSiteModel.PolSourceSiteTVItemID
                                 select c.IsActive).FirstOrDefault();

                if (ActivePollutionSource != IsActive)
                {
                    continue;
                }

                countPolSourceSite += 1;
                PolSourceObservationModel polSourceObservationModel = polSourceSiteService._PolSourceObservationService.GetPolSourceObservationModelLatestWithPolSourceSiteTVItemIDDB(polSourceSiteModel.PolSourceSiteTVItemID);

                List<MapInfoPointModel> mapInfoPointModelList = polSourceSiteService._MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                row = XlsxBase.AddRow();

                if (countPolSourceSite % 5 == 0)
                {
                    XlsxBase.CurrentBorderStyleValue = BorderStyleValues.Thin;
                    XlsxBase.CurrentBottomBorder = true;
                }
                else
                {
                    XlsxBase.CurrentBorderStyleValue = null;
                    XlsxBase.CurrentBottomBorder = false;
                }

                XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;
                cell = XlsxBase.AddCellString(row, polSourceSiteModel.Site.ToString());

                if (polSourceObservationModel != null)
                {
                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = polSourceSiteService._PolSourceObservationService._PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithPolSourceObservationIDDB(polSourceObservationModel.PolSourceObservationID);

                    string SelectedObservation = "Selected: \r\n";
                    int PolSourceObsInfoInt = 0;
                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList)
                    {
                        List<int> observationInfoList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList().Select(c => int.Parse(c)).ToList();
                        PolSourceObsInfoInt = observationInfoList.Where(c => (c > 10500 && c < 10599) || (c > 15200 && c < 15299)).FirstOrDefault();

                        foreach (PolSourceObsInfoEnum polSourceObsInfo in polSourceObservationIssueModel.PolSourceObsInfoList)
                        {
                            SelectedObservation += baseEnumService.GetEnumText_PolSourceObsInfoReportEnum(polSourceObsInfo);
                        }
                        SelectedObservation += "\r\n\r\n";
                    }
                    if (PolSourceObsInfoInt > 0)
                    {
                        cell = XlsxBase.AddCellString(row, baseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)PolSourceObsInfoInt));
                    }
                    else
                    {
                        cell = XlsxBase.AddCellString(row, "Empty");
                    }
                }

                XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Left;
                XlsxBase.WrapText = true;

                if (polSourceObservationModel != null)
                {
                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = polSourceSiteService._PolSourceObservationService._PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithPolSourceObservationIDDB(polSourceObservationModel.PolSourceObservationID);

                    string SelectedObservation = "Selected: \r\n";
                    int PolSourceObsInfoInt = 0;
                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList)
                    {
                        List<int> observationInfoList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList().Select(c => int.Parse(c)).ToList();
                        foreach (PolSourceObsInfoEnum polSourceObsInfo in polSourceObservationIssueModel.PolSourceObsInfoList)
                        {
                            SelectedObservation += baseEnumService.GetEnumText_PolSourceObsInfoReportEnum(polSourceObsInfo);
                        }
                        SelectedObservation += "\r\n\r\n";
                    }
                    if (PolSourceObsInfoInt > 0)
                    {
                        cell = XlsxBase.AddCellString(row, "Written: \r\n" + (string.IsNullOrWhiteSpace(polSourceObservationModel.Observation_ToBeDeleted) ? "" : polSourceObservationModel.Observation_ToBeDeleted.ToString()) + "\r\n\r\n"
                           + SelectedObservation);
                    }
                    else
                    {
                        cell = XlsxBase.AddCellString(row, "Written: \r\n\r\n" + SelectedObservation);
                    }
                }

                XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;
                Alignment alignment1 = new Alignment() { WrapText = true };

                if (mapInfoPointModelList.Count > 0)
                {
                    cell = XlsxBase.AddCellString(row, mapInfoPointModelList[0].Lat.ToString("F5"));
                    cell = XlsxBase.AddCellString(row, mapInfoPointModelList[0].Lng.ToString("F5"));
                }
                else
                {
                    cell = XlsxBase.AddCellString(row, "");
                    cell = XlsxBase.AddCellString(row, "");
                }

                TVItemModel tvItemModelPolSource = tvItemService.GetTVItemModelWithTVItemIDDB(polSourceSiteModel.PolSourceSiteTVItemID);

                cell = XlsxBase.AddCellString(row, tvItemModelPolSource.IsActive.ToString());
                cell = XlsxBase.AddCellString(row, polSourceObservationModel.ObservationDate_Local.ToString("yyyy MMM dd"));
                cell = XlsxBase.AddCellString(row, "empty");

                sheetData.Append(row);
            }

            if (countPolSourceSite > 0)
            {
                DefinedNames definedNames1 = new DefinedNames();
                DefinedName definedName1 = new DefinedName() { Name = "_xlnm._FilterDatabase", LocalSheetId = (UInt32Value)0U, Hidden = true };
                definedName1.Text = "'" + SheetName + "'" + "!$A$2:$H$" + (countPolSourceSite + 2).ToString();

                definedNames1.Append(definedName1);
                workbook.Append(definedNames1);

                AutoFilter autoFilter = new AutoFilter() { Reference = "a2:B" + (countPolSourceSite + 2).ToString() };
                worksheet.Append(autoFilter);
            }

            PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup = new PageSetup() { Orientation = OrientationValues.Portrait, Id = "rId" + SheetOrdinal.ToString() };

            worksheet.Append(sheetViews);
            worksheet.Append(sheetFormatProperties);

            if (columns.ChildElements.Count > 0)
                worksheet.Append(columns);

            worksheet.Append(sheetData);

            mergeCells.Count = (UInt32Value)((UInt32)mergeCells.ChildElements.Count);
            if (mergeCells.ChildElements.Count > 0)
                worksheet.Append(mergeCells);


            if (XlsxBase.UsedHyperlinkList.Count > 0)
                worksheet.Append(hyperlinks);

            worksheet.Append(pageMargins);
            worksheet.Append(pageSetup);

            worksheetPart.Worksheet = worksheet;

        }
    }
}