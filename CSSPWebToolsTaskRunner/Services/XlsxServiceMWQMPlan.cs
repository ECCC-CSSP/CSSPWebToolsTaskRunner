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
    public class XlsxServiceMWQMPlan
    {
        #region Variables
        #endregion Variables

        #region Properties
        public GeneratedClassMWQMPlanXlsx _Xlsx { get; set; }
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public XlsxServiceMWQMPlan(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _Xlsx = new GeneratedClassMWQMPlanXlsx(_TaskRunnerBaseService);
        }
        #endregion Constructors

        public void Generate(FileInfo fi)
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
            _Xlsx.CreatePackage(fi.FullName);
        }
    }

    public class GeneratedClassMWQMPlanXlsx
    {
        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public XlsxServiceBase XlsxBase { get; private set; }
        #endregion Properties

        #region Constructors
        public GeneratedClassMWQMPlanXlsx(TaskRunnerBaseService taskRunnerBaseService)
        {
            XlsxBase = new XlsxServiceBase(taskRunnerBaseService);
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions
        #endregion Functions

        public void CreatePackage(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }
        }
        private void CreateParts(SpreadsheetDocument document)
        {
            int SheetOrdinal = 0;

            BaseEnumService baseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            XlsxBase.AddSheetNameAndID(baseEnumService.GetEnumText_TVTypeEnum(TVTypeEnum.MWQMPlan));

            ExtendedFilePropertiesPart extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>("rId1000");
            XlsxBase.GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart);
        
            WorkbookPart workbookPart = document.AddWorkbookPart();
            XlsxBase.GenerateWorkbookPart1Content(workbookPart);

            SheetOrdinal = 0;
            WorksheetPart worksheetA = workbookPart.AddNewPart<WorksheetPart>(XlsxBase.sheetNameAndIDList[SheetOrdinal].SheetID);
            GenerateWorksheetAContent(worksheetA, SheetOrdinal);

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
        private void GenerateWorksheetAContent(WorksheetPart worksheetPart, int SheetOrdinal)
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
            List<double?> columnWidthList = new List<double?>() { 20D, 20D, 20D };
            foreach (double? width in columnWidthList)
            {
                Column colum = XlsxBase.AddColumnProp(columns, width);
            }

            SheetData sheetData = new SheetData();

            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMPlanService mwqmPlanService = new MWQMPlanService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMPlanSubsectorService mwqmPlanSubsectorService = new MWQMPlanSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMPlanSubsectorSiteService mwqmPlanSubsectorSiteService = new MWQMPlanSubsectorSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            MWQMPlanModel mwqmPlanModel = mwqmPlanService.GetMWQMPlanModelWithMWQMPlanIDDB(_TaskRunnerBaseService._BWObj.MWQMPlanID);

            XlsxBase.CurrentRowHeight = 18D;
            XlsxBase.CurrentFontSize = 16;
            XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

            if (!string.IsNullOrWhiteSpace(mwqmPlanModel.Error))
            {
                // Config File Name
                row = XlsxBase.AddRow();
                XlsxBase.CurrentFontColor = null;
                cell = XlsxBase.AddCellString(row, XslxServiceMWQMPlanRes.Error + " [" + mwqmPlanModel.Error + "]");
                sheetData.Append(row);
            }
            else
            {
                XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

                // Province and Year
                string Province = tvItemService.GetTVItemModelWithTVItemIDDB(mwqmPlanModel.ProvinceTVItemID).TVText;
                row = XlsxBase.AddRow();
                XlsxBase.CurrentFontColor = null;
                cell = XlsxBase.AddCellString(row, Province + "    (" + mwqmPlanModel.Year.ToString() + ")");
                cell = XlsxBase.AddCellString(row, "");
                cell = XlsxBase.AddCellString(row, "");
                sheetData.Append(row);

                MergeCell mergeCell = new MergeCell() { Reference = "A" + XlsxBase.CurrentRow.ToString() + ":C" + XlsxBase.CurrentRow.ToString() };

                mergeCells.Append(mergeCell);

                // Config File Name
                row = XlsxBase.AddRow();
                XlsxBase.CurrentFontColor = null;
                cell = XlsxBase.AddCellString(row, XslxServiceMWQMPlanRes.FileName + "   (config_" + mwqmPlanModel.ConfigFileName + ".txt)");
                cell = XlsxBase.AddCellString(row, "");
                cell = XlsxBase.AddCellString(row, "");
                sheetData.Append(row);

                mergeCell = new MergeCell() { Reference = "A" + XlsxBase.CurrentRow.ToString() + ":C" + XlsxBase.CurrentRow.ToString() };

                mergeCells.Append(mergeCell);

                // For Group Name
                row = XlsxBase.AddRow();
                XlsxBase.CurrentFontColor = null;
                cell = XlsxBase.AddCellString(row, XslxServiceMWQMPlanRes.For + "   (" + mwqmPlanModel.ForGroupName + ")");
                cell = XlsxBase.AddCellString(row, "");
                cell = XlsxBase.AddCellString(row, "");
                sheetData.Append(row);

                mergeCell = new MergeCell() { Reference = "A" + XlsxBase.CurrentRow.ToString() + ":C" + XlsxBase.CurrentRow.ToString() };

                mergeCells.Append(mergeCell);

               
                row = XlsxBase.AddRow();
                sheetData.Append(row);

                List<MWQMPlanSubsectorModel> mwqmPlanSubsectorModelList = mwqmPlanSubsectorService.GetMWQMPlanSubsectorModelListWithMWQMPlanIDDB(mwqmPlanModel.MWQMPlanID);

                foreach (MWQMPlanSubsectorModel mwqmPlanSubsectorModel in mwqmPlanSubsectorModelList)
                {
                    XlsxBase.CurrentRowHeight = 18D;
                    XlsxBase.CurrentFontSize = 16;
                    XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

                    row = XlsxBase.AddRow();
                    string Subsector = tvItemService.GetTVItemModelWithTVItemIDDB(mwqmPlanSubsectorModel.SubsectorTVItemID).TVText;
                    XlsxBase.CurrentFontColor = null;
                    cell = XlsxBase.AddCellString(row, Subsector);
                    cell = XlsxBase.AddCellString(row, "");
                    cell = XlsxBase.AddCellString(row, "");
                    sheetData.Append(row);

                    mergeCell = new MergeCell() { Reference = "A" + XlsxBase.CurrentRow.ToString() + ":C" + XlsxBase.CurrentRow.ToString() };

                    mergeCells.Append(mergeCell);

                    XlsxBase.CurrentRowHeight = 16D;
                    XlsxBase.CurrentFontSize = 12;
                    XlsxBase.CurrentBorderStyleValue = BorderStyleValues.Thin;
                    XlsxBase.CurrentBottomBorder = true;
                    XlsxBase.CurrentBorderColor = System.Drawing.Color.Green;
                    XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

                    List<MWQMPlanSubsectorSiteModel> mwqmPlanSubsectorSiteModelList = mwqmPlanSubsectorSiteService.GetMWQMPlanSubsectorSiteModelListWithMWQMPlanSubsectorIDDB(mwqmPlanSubsectorModel.MWQMPlanSubsectorID);

                    row = XlsxBase.AddRow();
                    cell = XlsxBase.AddCellString(row, XslxServiceMWQMPlanRes.MWQMSite);
                    cell = XlsxBase.AddCellString(row, XslxServiceMWQMPlanRes.Lat);
                    cell = XlsxBase.AddCellString(row, XslxServiceMWQMPlanRes.Lng);
                    sheetData.Append(row);

                    int count = 0;
                    foreach (MWQMPlanSubsectorSiteModel mwqmPlanSubsectorSiteModel in mwqmPlanSubsectorSiteModelList)
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
                        cell = XlsxBase.AddCellString(row, mwqmPlanSubsectorSiteModel.MWQMSiteText);

                        List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mwqmPlanSubsectorSiteModel.MWQMSiteTVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);
                        XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;
                        if (mapInfoPointModelList.Count > 0)
                        {
                            cell = XlsxBase.AddCellNumber(row, mapInfoPointModelList[0].Lat.ToString("F5"));
                            cell = XlsxBase.AddCellNumber(row, mapInfoPointModelList[0].Lng.ToString("F5"));
                        }
                        else
                        {
                            cell = XlsxBase.AddCellNumber(row, "");
                            cell = XlsxBase.AddCellNumber(row, "");
                        }
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