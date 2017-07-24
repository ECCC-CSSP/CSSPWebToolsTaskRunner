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
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Services;

namespace CSSPWebToolsTaskRunner.Services
{
    public class XlsxServiceSector
    {
        #region Variables
        #endregion Variables

        #region Properties
        public GeneratedClassSectorXlsx _Xlsx { get; set; }
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public XlsxServiceSector(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _Xlsx = new GeneratedClassSectorXlsx(_TaskRunnerBaseService);
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

    public class GeneratedClassSectorXlsx
    {
        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public XlsxServiceBase XlsxBase { get; private set; }
        #endregion Properties

        #region Constructors
        public GeneratedClassSectorXlsx(TaskRunnerBaseService taskRunnerBaseService)
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
            XlsxBase.AddSheetNameAndID(baseEnumService.GetEnumText_TVTypeEnum(TVTypeEnum.Sector));

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
            List<double?> columnWidthList = new List<double?>() { 30D };
            foreach (double? width in columnWidthList)
            {
                Column colum = XlsxBase.AddColumnProp(columns, width);
            }

            SheetData sheetData = new SheetData();

            BaseEnumService baseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemStatService tvItemStatService = new TVItemStatService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            TVItemModel tvItemModelSector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSector.TVItemID, TVTypeEnum.Subsector);

        
            foreach (TVItemModel tvItemModelSubsector in tvItemModelSubsectorList)
            {
                //tvItemStatService.SetTVItemStatForTVItemIDAndParentsTVItemID(tvItemModelSubsector.TVItemID);
                
                XlsxBase.CurrentRowHeight = 24D;
                XlsxBase.CurrentFontSize = 18;
                XlsxBase.CurrentBorderStyleValue = BorderStyleValues.Thick;
                XlsxBase.CurrentBottomBorder = true;
                XlsxBase.CurrentBorderColor = System.Drawing.Color.Green;
                XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

                row = XlsxBase.AddRow();
                string URL = _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelSubsector);
                XlsxBase.CurrentFontColor = System.Drawing.Color.Blue;
                cell = XlsxBase.AddCellHyperlink(hyperlinks, row, tvItemModelSubsector.TVText, URL);
                XlsxBase.CurrentFontColor = null;
                //cell = XlsxBase.AddCellString(row, tvItemModel.TVText);
                cell = XlsxBase.AddCellString(row, null);
                sheetData.Append(row);

                MergeCell mergeCell = new MergeCell() { Reference = "A" + XlsxBase.CurrentRow.ToString() + ":B" + XlsxBase.CurrentRow.ToString() };

                mergeCells.Append(mergeCell);

                XlsxBase.CurrentRowHeight = 16D;
                XlsxBase.CurrentFontSize = 12;
                XlsxBase.CurrentHorizontalAlignmentValue = HorizontalAlignmentValues.Center;

                List<TVItemStatModel> tvItemStatModelList = tvItemStatService.GetTVItemStatModelListWithTVItemIDDB(tvItemModelSubsector.TVItemID);

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

    }
}