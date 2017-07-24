using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public class XlsxServiceBase
    {
        #region Variables
        #endregion Variables

        #region Properties
        TaskRunnerBaseService _TaskRunnerBaseService { get; set; }
        public List<IndexFont> indexFontList { get; set; }
        public List<IndexNumberingFormat> indexNumberingFormatList { get; set; }
        public List<IndexFill> indexFillList { get; set; }
        public List<IndexBorder> indexBorderList { get; set; }
        public List<IndexCellFormat> indexCellFormatList { get; set; }
        public List<IndexCellStyleFormat> indexCellStyleFormatList { get; set; }
        public List<IndexDifferentialFormat> indexDifferentialFormatList { get; set; }
        public List<IndexCellStyle> indexCellStyleList { get; set; }
        public List<IndexSharedStringItem> indexSharedStringItemList { get; set; }
        public List<IndexFormat> indexFormatList { get; set; }
        public int CurrentID { get; set; }
        public List<SheetNameAndID> sheetNameAndIDList { get; set; }
        public string LastCell { get; set; }
        public UInt32 CurrentRow { get; set; }
        public UInt32 CurrentColumn { get; set; }
        public UInt32 CurrentColumnProp { get; set; }
        public double CurrentRowHeight { get; set; }
        public int CurrentFontSize { get; set; }
        public System.Drawing.Color? CurrentFontColor { get; set; }
        public FontNameEnum? CurrentFontName { get; set; }
        public bool CurrentFontBold { get; set; }
        public bool CurrentFontItalic { get; set; }
        public bool CurrentFontUnderline { get; set; }
        public PatternValues? CurrentPatternValue { get; set; }
        public System.Drawing.Color? CurrentPatternBGColor { get; set; }
        public System.Drawing.Color? CurrentPatternFGColor { get; set; }
        public BorderStyleValues? CurrentBorderStyleValue { get; set; }
        public bool CurrentLeftBorder { get; set; }
        public bool CurrentRightBorder { get; set; }
        public bool CurrentTopBorder { get; set; }
        public bool CurrentBottomBorder { get; set; }
        public bool CurrentDiagonalBorder { get; set; }
        public System.Drawing.Color? CurrentBorderColor { get; set; }
        public string CurrentFormatCode { get; set; }
        public bool? WrapText { get; set; }
        public HorizontalAlignmentValues? CurrentHorizontalAlignmentValue { get; set; }
        public VerticalAlignmentValues? CurrentVerticalAlignmentValue { get; set; }
        public List<UsedHyperlink> UsedHyperlinkList { get; set; }
        #endregion Properties

        #region Constructors
        public XlsxServiceBase(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            indexFontList = new List<IndexFont>();
            indexNumberingFormatList = new List<IndexNumberingFormat>();
            indexFillList = new List<IndexFill>();
            indexBorderList = new List<IndexBorder>();
            indexCellFormatList = new List<IndexCellFormat>();
            indexCellStyleFormatList = new List<IndexCellStyleFormat>();
            indexDifferentialFormatList = new List<IndexDifferentialFormat>();
            indexCellStyleList = new List<IndexCellStyle>();
            indexSharedStringItemList = new List<IndexSharedStringItem>();
            indexFormatList = new List<IndexFormat>();
            sheetNameAndIDList = new List<SheetNameAndID>();
            ClearAllList();
            UsedHyperlinkList = new List<UsedHyperlink>();
        }
        #endregion Constructors

        #region Functions public
        public Cell AddCellFormat()
        {
            UInt32 NumberingFormatInt = 0U;
            if (CurrentFormatCode.Length > 0)
                NumberingFormatInt = SetNumberingFormat();

            UInt32 FontInt = 0U;
            FontInt = SetFont();

            UInt32 FillInt = 0U;
            FillInt = SetFill();

            UInt32 BorderInt = 0U;
            BorderInt = SetBorder();

            UInt32 CellStyleFormatInt = 0U;
            CellStyleFormatInt = SetCellStyleFormat(0U, 0U, 0U, 0U, 0U);

            UInt32 CellFormatInt = 0U;
            UInt32 FormatId2 = 0U;
            CellFormatInt = SetCellFormat(NumberingFormatInt, FontInt, FillInt, BorderInt, FormatId2);

            Cell cell = new Cell() { CellReference = GetCellReference(CurrentColumn, CurrentRow) };

            cell.StyleIndex = (UInt32Value)CellFormatInt;

            return cell;
        }
        public Cell AddCellFormula(Row row, FormulaEnum Command, string CellArea, string CellText)
        {
            CurrentColumn += 1;

            Cell cell = AddCellFormat();

            if ((int)Command > 0 && CellArea.Length > 0)
            {
                CellFormula cellFormula = new CellFormula();
                if (Command == FormulaEnum.SUM)
                {
                    cellFormula.Text = Command + "(" + CellArea + ")";
                }
                else if (Command == FormulaEnum.AVERAGE)
                {
                    cellFormula.Text = Command + "(" + CellArea + ")";
                }
                else
                {
                    cellFormula.Text = CellArea;
                }
                CellValue cellValue = new CellValue();
                cellValue.Text = CellText;

                cell.Append(cellFormula);
                cell.Append(cellValue);
            }
            else
            {
            }

            row.Append(cell);

            return cell;
        }
        public Cell AddCellHyperlink(Hyperlinks hyperlinks, Row row, string SharedString, string URL)
        {
            CurrentColumn += 1;
            if (SharedString == null)
            {
            }
            else if (SharedString.Length == 0)
            {
                return null;
            }

            Cell cell = AddCellFormat();
            if (SharedString != null && SharedString.Length > 0)
            {
                cell.DataType = CellValues.SharedString;
                UInt32 SharedInt = SetSharedStringItem(SharedString);
                CellValue cellValue = new CellValue();
                cellValue.Text = SharedInt.ToString();
                cell.Append(cellValue);
            }
            else
            {
            }

            UsedHyperlink usedHyperlink = (from c in UsedHyperlinkList
                                           where c.URL == URL
                                           select c).FirstOrDefault();

            Hyperlink hyperlink = new Hyperlink();
            if (usedHyperlink == null)
            {
                string Id = "rId" + (100 + UsedHyperlinkList.Count).ToString();
                hyperlink.Id = Id;

                usedHyperlink = new UsedHyperlink()
                {
                    URL = URL,
                    Id = Id,
                };

                UsedHyperlinkList.Add(usedHyperlink);
            }
            else
            {
                hyperlink.Id = usedHyperlink.Id;
            }

            hyperlink.Reference = GetCellReference(CurrentColumn, CurrentRow);
            hyperlinks.Append(hyperlink);

            row.Append(cell);

            return cell;
        }
        public Cell AddCellNumber(Row row, string NumberText)
        {
            float TempFloat = -99999999999.0f;
            float.TryParse(NumberText, out TempFloat);

            if (TempFloat == -99999999999.0f)
                return null;

            CurrentColumn += 1;
            if (NumberText == null)
            {
            }
            else if (NumberText.Length == 0)
            {
                return null;
            }

            Cell cell = AddCellFormat();

            if (NumberText != null && NumberText.Length > 0)
            {
                CellValue cellValue = new CellValue();
                cellValue.Text = NumberText;
                cell.Append(cellValue);
            }
            else
            {
            }

            row.Append(cell);

            return cell;
        }
        public Cell AddCellString(Row row, string SharedString)
        {
            CurrentColumn += 1;
            if (SharedString == null)
            {
            }
            else if (SharedString.Length == 0)
            {
                return null;
            }

            Cell cell = AddCellFormat();
            if (SharedString != null && SharedString.Length > 0)
            {
                cell.DataType = CellValues.SharedString;
                UInt32 SharedInt = SetSharedStringItem(SharedString);
                CellValue cellValue = new CellValue();
                cellValue.Text = SharedInt.ToString();
                cell.Append(cellValue);
            }
            else
            {
            }

            row.Append(cell);

            return cell;
        }
        public Column AddColumnProp(Columns columns, double? Width)
        {
            CurrentColumnProp += 1;
            Column column = new Column() { Min = (UInt32Value)CurrentColumnProp, Max = (UInt32Value)CurrentColumnProp, Width = Width, CustomWidth = true };

            if (Width != null)
                columns.Append(column);

            return column;
        }
        public Row AddRow()
        {
            CurrentColumn = 0;
            CurrentRow += 1;
            Row row = new Row() { RowIndex = (UInt32Value)CurrentRow, Spans = new ListValue<StringValue>() { InnerText = "1:1" }, Height = CurrentRowHeight, DyDescent = 0.25D };

            return row;
        }
        public void AddSheetNameAndID(string SheetName)
        {
            CurrentID += 1;
            sheetNameAndIDList.Add(new SheetNameAndID() { SheetName = SheetName, SheetID = "rId" + CurrentID.ToString() });
        }
        public void ClearAllList()
        {
            CurrentColumnProp = 0;
            CurrentRow = 0;
            CurrentColumn = 0;
            CurrentID = 0;
            indexFontList.Clear();
            indexNumberingFormatList.Clear();
            indexFillList.Clear();
            indexBorderList.Clear();
            indexCellFormatList.Clear();
            indexDifferentialFormatList.Clear();
            indexCellStyleList.Clear();
            indexSharedStringItemList.Clear();
            sheetNameAndIDList.Clear();

            CurrentRowHeight = 18D;

            CurrentFontSize = 12;
            CurrentFontColor = null;
            CurrentFontName = FontNameEnum.Arial;
            CurrentFontBold = false;
            CurrentFontItalic = false;
            CurrentFontUnderline = false;
            SetFont();

            CurrentFormatCode = "";
            SetNumberingFormat();

            CurrentPatternValue = PatternValues.None;
            CurrentPatternBGColor = null;
            CurrentPatternFGColor = null;
            SetFill();

            CurrentPatternValue = PatternValues.Gray125;
            SetFill();

            CurrentPatternValue = PatternValues.None;

            CurrentBorderStyleValue = null;
            CurrentLeftBorder = false;
            CurrentRightBorder = false;
            CurrentTopBorder = false;
            CurrentBottomBorder = false;
            CurrentDiagonalBorder = false;
            CurrentBorderColor = null;
            SetBorder();

            SetCellStyleFormat(0U, 0U, 0U, 0U, 0U);

            WrapText = null;
            CurrentHorizontalAlignmentValue = null;
            CurrentVerticalAlignmentValue = null;
            SetCellFormat(0U, 0U, 0U, 0U, 0U);

            SetCellStyle("Normal", 0U, 0U);

            // doing hyperlink
            CurrentFontColor = System.Drawing.Color.Blue;
            SetFont();
            SetCellFormat(0U, 1U, 0U, 0U, 1U);
            SetCellStyleFormat(0U, 1U, 0U, 0U, 1U);
            SetCellStyle("Hyperlink", 1U, 8U);

            CurrentFontColor = null;
        }
        public void GenerateCalculationChainPart1Content(CalculationChainPart calculationChainPart)
        {
            CalculationChain calculationChain = new CalculationChain();
            CalculationCell calculationCell = new CalculationCell() { CellReference = "A3", SheetId = 2, NewLevel = true };

            calculationChain.Append(calculationCell);

            calculationChainPart.CalculationChain = calculationChain;
        }
        public void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "3";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            UInt32Value Size = UInt32Value.FromUInt32((System.UInt32)sheetNameAndIDList.Count);
            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = Size };
            foreach (SheetNameAndID sheetNameAndID in sheetNameAndIDList)
            {
                Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
                vTLPSTR2.Text = sheetNameAndID.SheetName;

                vTVector2.Append(vTLPSTR2);
            }

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "Environment Canada";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart.Properties = properties1;
        }
        public void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart)
        {
            UInt32 Count = (UInt32)indexSharedStringItemList.Count();
            UInt32 UniqueCount = (UInt32)indexSharedStringItemList.Count();
            SharedStringTable sharedStringTable = new SharedStringTable() { Count = (UInt32Value)Count, UniqueCount = (UInt32Value)UniqueCount };

            foreach (IndexSharedStringItem indexSharedStringItem in indexSharedStringItemList)
            {
                SharedStringItem sharedStringItem = new SharedStringItem();
                Text text = new Text();
                text.Text = indexSharedStringItem.CellText;

                sharedStringItem.Append(text);

                sharedStringTable.Append(sharedStringItem);
            }

            sharedStringTablePart.SharedStringTable = sharedStringTable;
        }
        public void GenerateWorkbookPart1Content(WorkbookPart workbookPart)
        {
            Workbook workbook = new Workbook();
            workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion = new FileVersion() { ApplicationName = "xl", LastEdited = "5", LowestEdited = "5", BuildVersion = "9303" };
            WorkbookProperties workbookProperties = new WorkbookProperties();

            BookViews bookViews = new BookViews();
            WorkbookView workbookView = new WorkbookView() { XWindow = 270, YWindow = 630, WindowWidth = (UInt32Value)24735U, WindowHeight = (UInt32Value)11445U };

            bookViews.Append(workbookView);

            Sheets sheets = new Sheets();
            for (int i = 0; i < sheetNameAndIDList.Count; i++)
            {
                Sheet sheet = new Sheet() { Name = sheetNameAndIDList[i].SheetName, SheetId = (UInt32Value)(i + 1U), Id = sheetNameAndIDList[i].SheetID };

                sheets.Append(sheet);
            }

            CalculationProperties calculationProperties = new CalculationProperties() { CalculationId = (UInt32Value)145621U };

            workbook.Append(fileVersion);
            workbook.Append(workbookProperties);
            workbook.Append(bookViews);
            workbook.Append(sheets);
            workbook.Append(calculationProperties);

            workbookPart.Workbook = workbook;
        }
        public void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            UInt32 FontsCount = (UInt32)indexFontList.Count;
            Fonts fonts = new Fonts() { Count = (UInt32Value)FontsCount, KnownFonts = true };
            foreach (IndexFont indexFont in indexFontList)
            {
                fonts.Append(indexFont.Font);
            }

            UInt32 FillsCount = (UInt32)indexFillList.Count;
            Fills fills = new Fills() { Count = (UInt32Value)FillsCount };
            foreach (IndexFill indexFill in indexFillList)
            {
                fills.Append(indexFill.Fill);
            }

            UInt32 BordersCount = (UInt32)indexBorderList.Count;
            Borders borders = new Borders() { Count = (UInt32Value)BordersCount };
            foreach (IndexBorder indexBorder in indexBorderList)
            {
                borders.Append(indexBorder.Border);
            }

            UInt32 CellStyleFormatCount = (UInt32)indexCellStyleFormatList.Count();
            CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = (UInt32Value)CellStyleFormatCount };
            foreach (IndexCellStyleFormat indexCellStyleFormat in indexCellStyleFormatList)
            {
                cellStyleFormats.Append(indexCellStyleFormat.CellFormat);
            }

            UInt32 CellFormatsCount = (UInt32)indexCellFormatList.Count();
            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)CellFormatsCount };
            foreach (IndexCellFormat indexCellFormat in indexCellFormatList)
            {
                cellFormats.Append(indexCellFormat.CellFormat);
            }

            UInt32 CellStyleCount = (UInt32)indexCellStyleList.Count;
            CellStyles cellStyles = new CellStyles() { Count = (UInt32Value)CellStyleCount };
            foreach (IndexCellStyle indexCellStyle in indexCellStyleList)
            {
                cellStyles.Append(indexCellStyle.CellStyle);
            }

            DifferentialFormats differentialFormats = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension.Append(slicerStyles);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles);

            stylesheetExtensionList.Append(stylesheetExtension);
            stylesheetExtensionList.Append(stylesheetExtension2);

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles);
            stylesheet.Append(differentialFormats);
            stylesheet.Append(tableStyles);
            stylesheet.Append(stylesheetExtensionList);

            workbookStylesPart.Stylesheet = stylesheet;
        }
        public void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }
        public string GetCellReference(UInt32 column, UInt32 row)
        {
            if (column < 1 || row < 1)
            {
                return "";
            }
            else if (column > 0 && column <= 26)
            {
                return ((char)(64 + column)).ToString() + row.ToString();
            }
            else if (column > 26 && column <= 52)
            {
                return "A" + ((char)(64 + column)).ToString() + row.ToString();
            }
            else if (column > 52 && column <= 78)
            {
                return "B" + ((char)(64 + column)).ToString() + row.ToString();
            }
            else
            {
                return "";
            }
        }
        public string GetFontName(FontNameEnum? fontNameEnum)
        {
            if (fontNameEnum == null)
                return "Arial";

            switch (fontNameEnum)
            {
                case FontNameEnum.Arial:
                    return "Arial";
                case FontNameEnum.CourrierNew:
                    return "Courrier New";
                default:
                    return "Arial";
            }
        }
        public UInt32 SetBorder()
        {
            string bRGB = null;
            if (CurrentBorderColor != null)
            {
                bRGB = CurrentBorderColor.Value.A.ToString("X2") + CurrentBorderColor.Value.R.ToString("X2") + CurrentBorderColor.Value.G.ToString("X2") + CurrentBorderColor.Value.B.ToString("X2");
            }

            Border border = new Border();
            LeftBorder LeftBorder1 = new LeftBorder();
            if (CurrentLeftBorder && CurrentBorderStyleValue != null)
            {
                LeftBorder1.Style = CurrentBorderStyleValue;
            }
            if (CurrentLeftBorder && !string.IsNullOrWhiteSpace(bRGB))
            {
                LeftBorder1.Color = new Color() { Rgb = bRGB };
            }
            border.Append(LeftBorder1);
            RightBorder RightBorder1 = new RightBorder();
            if (CurrentRightBorder && CurrentBorderStyleValue != null)
            {
                RightBorder1.Style = CurrentBorderStyleValue;
            }
            if (CurrentRightBorder && !string.IsNullOrWhiteSpace(bRGB))
            {
                RightBorder1.Color = new Color() { Rgb = bRGB };
            }
            border.Append(RightBorder1);
            TopBorder TopBorder1 = new TopBorder();
            if (CurrentTopBorder && CurrentBorderStyleValue != null)
            {
                TopBorder1.Style = CurrentBorderStyleValue;
            }
            if (CurrentTopBorder && !string.IsNullOrWhiteSpace(bRGB))
            {
                TopBorder1.Color = new Color() { Rgb = bRGB };
            }
            border.Append(TopBorder1);
            BottomBorder BottomBorder1 = new BottomBorder();
            if (CurrentBottomBorder && CurrentBorderStyleValue != null)
            {
                BottomBorder1.Style = CurrentBorderStyleValue;
            }
            if (CurrentBottomBorder && !string.IsNullOrWhiteSpace(bRGB))
            {
                BottomBorder1.Color = new Color() { Rgb = bRGB };
            }
            border.Append(BottomBorder1);
            DiagonalBorder DiagonalBorder1 = new DiagonalBorder();
            if (CurrentDiagonalBorder && CurrentBorderStyleValue != null)
            {
                DiagonalBorder1.Style = CurrentBorderStyleValue;
            }
            if (CurrentDiagonalBorder && !string.IsNullOrWhiteSpace(bRGB))
            {
                DiagonalBorder1.Color = new Color() { Rgb = bRGB };
            }
            border.Append(DiagonalBorder1);

            IndexBorder indexBorder = (from c in indexBorderList
                                       where c.borderStyleValues == CurrentBorderStyleValue
                                           && c.leftBorder == CurrentLeftBorder
                                           && c.rightBorder == CurrentRightBorder
                                           && c.topBorder == CurrentTopBorder
                                           && c.bottomBorder == CurrentBottomBorder
                                           && c.diagonalBorder == CurrentDiagonalBorder
                                           && c.borderRGB == bRGB
                                       select c).FirstOrDefault();
            UInt32 Index = 0U;
            if (indexBorder == null)
            {
                Index = (UInt32)indexBorderList.Count;
                indexBorderList.Add(new IndexBorder()
                {
                    Index = Index,
                    borderStyleValues = CurrentBorderStyleValue,
                    leftBorder = CurrentLeftBorder,
                    rightBorder = CurrentRightBorder,
                    topBorder = CurrentTopBorder,
                    bottomBorder = CurrentBottomBorder,
                    diagonalBorder = CurrentDiagonalBorder,
                    borderRGB = bRGB,
                    Border = border
                });
            }
            else
            {
                Index = indexBorder.Index;
            }

            return Index;
        }
        public UInt32 SetCellFormat(UInt32 NumberFormatId, UInt32 FontId, UInt32 FillId, UInt32 BorderId, UInt32 FormatId)
        {
            CellFormat cellFormat = new CellFormat()
            {
                NumberFormatId = (UInt32Value)NumberFormatId,
                FontId = (UInt32Value)FontId,
                FillId = (UInt32Value)FillId,
                BorderId = (UInt32Value)BorderId,
                FormatId = (UInt32Value)FormatId,
            };
            if (CurrentHorizontalAlignmentValue != null || CurrentVerticalAlignmentValue != null || WrapText != null)
            {
                Alignment alignment = new Alignment();

                if (CurrentHorizontalAlignmentValue != null)
                    alignment.Horizontal = CurrentHorizontalAlignmentValue;

                if (CurrentVerticalAlignmentValue != null)
                    alignment.Vertical = CurrentVerticalAlignmentValue;

                if (WrapText != null)
                    alignment.WrapText = true;

                cellFormat.Alignment = alignment;
            }

            IndexCellFormat indexCellFormat = (from c in indexCellFormatList
                                               where c.NumberFormatId == NumberFormatId
                                               && c.FontId == FontId
                                               && c.FillId == FillId
                                               && c.BorderId == BorderId
                                               && c.FormatId == FormatId
                                               && c.WrapText == WrapText
                                               && c.horizontalAlignmentValue == CurrentHorizontalAlignmentValue
                                               && c.verticalAlignmentValue == CurrentVerticalAlignmentValue
                                               select c).FirstOrDefault();
            UInt32 Index = 0U;
            if (indexCellFormat == null)
            {
                Index = (UInt32)indexCellFormatList.Count;
                indexCellFormatList.Add(new IndexCellFormat()
                {
                    Index = Index,
                    NumberFormatId = NumberFormatId,
                    FontId = FontId,
                    FillId = FillId,
                    BorderId = BorderId,
                    FormatId = FormatId,
                    WrapText = WrapText,
                    horizontalAlignmentValue = CurrentHorizontalAlignmentValue,
                    verticalAlignmentValue = CurrentVerticalAlignmentValue,
                    CellFormat = cellFormat,
                });
            }
            else
            {
                Index = indexCellFormat.Index;
            }

            return Index;
        }
        public UInt32 SetCellStyle(string Name, UInt32 FormatId, UInt32 BuiltinId)
        {
            IndexCellStyle indexCellStyle = (from c in indexCellStyleList
                                             where c.Name == Name
                                             && c.FormatId == FormatId
                                             && c.BuildinId == BuiltinId
                                             select c).FirstOrDefault();

            CellStyle cellStyle = new CellStyle()
            {
                Name = Name,
                FormatId = FormatId,
                BuiltinId = BuiltinId,
            };

            UInt32 Index = 0U;
            if (indexCellStyle == null)
            {
                Index = (UInt32)indexCellStyleList.Count;
                indexCellStyleList.Add(new IndexCellStyle()
                {
                    Index = Index,
                    Name = Name,
                    FormatId = FormatId,
                    BuildinId = BuiltinId,
                    CellStyle = cellStyle
                });
            }
            else
            {
                Index = indexCellStyle.Index;
            }

            return Index;
        }
        public UInt32 SetCellStyleFormat(UInt32 NumberFormatId, UInt32 FontId, UInt32 FillId, UInt32 BorderId, UInt32 FormatId)
        {
            CellFormat cellFormat = new CellFormat()
            {
                NumberFormatId = (UInt32Value)NumberFormatId,
                FontId = (UInt32Value)FontId,
                FillId = (UInt32Value)FillId,
                BorderId = (UInt32Value)BorderId,
                FormatId = (UInt32Value)FormatId,
            };

            IndexCellStyleFormat indexCellStyleFormat = (from c in indexCellStyleFormatList
                                                         where c.NumberFormatId == NumberFormatId
                                                         && c.FontId == FontId
                                                         && c.FillId == FillId
                                                         && c.BorderId == BorderId
                                                         && c.FormatId == FormatId
                                                         select c).FirstOrDefault();
            UInt32 Index = 0U;
            if (indexCellStyleFormat == null)
            {
                Index = (UInt32)indexCellStyleFormatList.Count;
                indexCellStyleFormatList.Add(new IndexCellStyleFormat()
                {
                    Index = Index,
                    NumberFormatId = NumberFormatId,
                    FontId = FontId,
                    FillId = FillId,
                    BorderId = BorderId,
                    FormatId = FormatId,
                    CellFormat = cellFormat,
                });
            }
            else
            {
                Index = indexCellStyleFormat.Index;
            }

            return Index;
        }
        public UInt32 SetFill()
        {
            //FFFF0000
            string bgRGB = null;
            string fgRGB = null;
            if (CurrentPatternBGColor != null)
            {
                bgRGB = CurrentPatternBGColor.Value.A.ToString("X2") + CurrentPatternBGColor.Value.R.ToString("X2") + CurrentPatternBGColor.Value.G.ToString("X2") + CurrentPatternBGColor.Value.B.ToString("X2");
            }
            if (CurrentPatternFGColor != null)
            {
                fgRGB = CurrentPatternFGColor.Value.A.ToString("X2") + CurrentPatternFGColor.Value.R.ToString("X2") + CurrentPatternFGColor.Value.G.ToString("X2") + CurrentPatternFGColor.Value.B.ToString("X2");
            }

            Fill fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    PatternType = CurrentPatternValue
                },
            };
            if (!string.IsNullOrWhiteSpace(bgRGB))
            {
                fill.PatternFill.BackgroundColor = new BackgroundColor() { Rgb = bgRGB };
            }
            if (!string.IsNullOrWhiteSpace(fgRGB))
            {
                fill.PatternFill.ForegroundColor = new ForegroundColor() { Rgb = fgRGB };
            }

            IndexFill indexFill = (from c in indexFillList
                                   where c.PatternValue == CurrentPatternValue
                                   && c.bgRGB == bgRGB
                                   && c.fgRGB == fgRGB
                                   select c).FirstOrDefault();
            UInt32 Index = 0U;
            if (indexFill == null)
            {
                Index = (UInt32)indexFillList.Count;
                indexFillList.Add(new IndexFill()
                {
                    Index = Index,
                    PatternValue = CurrentPatternValue,
                    bgRGB = bgRGB,
                    fgRGB = fgRGB,
                    Fill = fill
                });
            }
            else
            {
                Index = indexFill.Index;
            }

            return Index;
        }
        public UInt32 SetFont()
        {
            string fontRGB = "";
            if (CurrentFontColor != null)
            {
                fontRGB = CurrentFontColor.Value.A.ToString("X2") + CurrentFontColor.Value.R.ToString("X2") + CurrentFontColor.Value.G.ToString("X2") + CurrentFontColor.Value.B.ToString("X2");
            }
            else
            {
                System.Drawing.Color TempColor = System.Drawing.Color.Black;
                fontRGB = TempColor.A.ToString("X2") + TempColor.R.ToString("X2") + TempColor.G.ToString("X2") + TempColor.B.ToString("X2");
            }

            Font font = new Font()
            {
                FontSize = new FontSize() { Val = CurrentFontSize },
                FontName = new FontName() { Val = GetFontName(CurrentFontName) },
                FontFamilyNumbering = new FontFamilyNumbering() { Val = (CurrentFontName == FontNameEnum.Arial ? 2 : 3) },
                Bold = (CurrentFontBold == true ? new Bold() : null),
                Italic = (CurrentFontItalic == true ? new Italic() : null),
                Underline = (CurrentFontUnderline == true ? new Underline() : null),
            };

            font.Color = new Color() { Rgb = fontRGB };

            IndexFont indexFont = (from c in indexFontList
                                   where c.FontSize == CurrentFontSize
                                   && c.FontRGB == fontRGB
                                   && c.FontName == GetFontName(CurrentFontName)
                                   && c.Bold == CurrentFontBold
                                   && c.Italic == CurrentFontItalic
                                   && c.Underline == CurrentFontUnderline
                                   select c).FirstOrDefault();
            UInt32 Index = 0U;
            if (indexFont == null)
            {
                Index = (UInt32)indexFontList.Count;
                indexFontList.Add(new IndexFont()
                {
                    Index = Index,
                    FontSize = CurrentFontSize,
                    FontRGB = fontRGB,
                    FontName = GetFontName(CurrentFontName),
                    Bold = CurrentFontBold,
                    Italic = CurrentFontItalic,
                    Underline = CurrentFontUnderline,
                    Font = font
                });
            }
            else
            {
                Index = indexFont.Index;
            }

            return Index;
        }
        public UInt32 SetFormat(Format format)
        {
            IndexFormat indexFormat = indexFormatList.Where(c => c.Format == format).FirstOrDefault();
            UInt32 Index = 0U;
            if (indexFormat == null)
            {
                Index = (UInt32)indexFormatList.Count;
                indexFormatList.Add(new IndexFormat() { Index = Index, Format = format });
            }
            else
            {
                Index = indexFormat.Index;
            }

            return Index;
        }
        public UInt32 SetNumberingFormat()
        {
            if (CurrentFormatCode.Length == 0)
                return 0U;

            IndexNumberingFormat indexNumberingFormat = (from c in indexNumberingFormatList
                                                         where c.FormatCode == CurrentFormatCode
                                                         select c).FirstOrDefault();
            UInt32 Index = 0;
            if (indexNumberingFormat == null)
            {
                Index = (UInt32)indexNumberingFormatList.Count;
                indexNumberingFormatList.Add(new IndexNumberingFormat()
                {
                    Index = Index,
                    FormatCode = CurrentFormatCode,
                    NumberingFormat = new NumberingFormat() { FormatCode = CurrentFormatCode },
                });
            }
            else
            {
                Index = indexNumberingFormat.Index;
            }

            return Index;
        }
        public void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "LeBlanc,Charles [Moncton]";
            document.PackageProperties.LastModifiedBy = "LeBlanc,Charles [Moncton]";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModelContact = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.LastUpdateContactTVItemID);
            if (string.IsNullOrWhiteSpace(tvItemModelContact.Error))
            {
                document.PackageProperties.Creator = tvItemModelContact.TVText;
                document.PackageProperties.LastModifiedBy = tvItemModelContact.TVText;
            }
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = DateTime.UtcNow;
            document.PackageProperties.Modified = DateTime.UtcNow;
        }
        public UInt32 SetSharedStringItem(string CellText)
        {
            IndexSharedStringItem indexSharedStringItem = (from c in indexSharedStringItemList
                                                           where c.CellText == CellText
                                                           select c).FirstOrDefault();
            UInt32 Index = 0U;
            if (indexSharedStringItem == null)
            {
                Index = (UInt32)indexSharedStringItemList.Count;
                indexSharedStringItemList.Add(new IndexSharedStringItem()
                {
                    Index = Index,
                    CellText = CellText,
                    SharedStringItem = new SharedStringItem() { Text = new Text() { Text = CellText } }
                });
            }
            else
            {
                Index = indexSharedStringItem.Index;
            }

            return Index;
        }

        #endregion Functions public


    }
}