using DocumentFormat.OpenXml.Spreadsheet;
using System;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;


namespace CSSPWebToolsTaskRunner.Services
{
    // Class Excel
    public class IndexFont
    {
        public IndexFont()
        {

        }

        public UInt32 Index { get; set; }
        public int FontSize { get; set; }
        public string FontRGB { get; set; }
        public string FontName { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public Font Font { get; set; }
    }
    public class IndexSharedStringItem
    {
        public IndexSharedStringItem()
        {
        }

        public UInt32 Index { get; set; }
        public string CellText { get; set; }
        public SharedStringItem SharedStringItem { get; set; }
    }
    public class IndexNumberingFormat
    {
        public IndexNumberingFormat()
        {
        }

        public UInt32 Index { get; set; }
        public string FormatCode { get; set; }
        public NumberingFormat NumberingFormat { get; set; }
    }
    public class IndexFill
    {
        public IndexFill()
        {
        }

        public UInt32 Index { get; set; }
        public PatternValues? PatternValue { get; set; }
        public string bgRGB { get; set; }
        public string fgRGB { get; set; }
        public Fill Fill { get; set; }
    }
    public class IndexBorder
    {
        public IndexBorder()
        {
        }

        public UInt32 Index { get; set; }
        public BorderStyleValues? borderStyleValues { get; set; }
        public bool topBorder { get; set; }
        public bool rightBorder { get; set; }
        public bool bottomBorder { get; set; }
        public bool leftBorder { get; set; }
        public bool diagonalBorder { get; set; }
        public string borderRGB { get; set; }
        public Border Border { get; set; }
    }
    public class IndexCellFormat
    {
        public IndexCellFormat()
        {
        }

        public UInt32 Index { get; set; }
        public UInt32 NumberFormatId { get; set; }
        public UInt32 FontId { get; set; }
        public UInt32 FillId { get; set; }
        public UInt32 BorderId { get; set; }
        public UInt32 FormatId { get; set; }
        public bool? WrapText { get; set; }
        public HorizontalAlignmentValues? horizontalAlignmentValue { get; set; }
        public VerticalAlignmentValues? verticalAlignmentValue { get; set; }
        public CellFormat CellFormat { get; set; }
    }
    public class IndexCellStyleFormat
    {
        public IndexCellStyleFormat()
        {
        }

        public UInt32 Index { get; set; }
        public UInt32 NumberFormatId { get; set; }
        public UInt32 FontId { get; set; }
        public UInt32 FillId { get; set; }
        public UInt32 BorderId { get; set; }
        public UInt32 FormatId { get; set; }
        public CellFormat CellFormat { get; set; }
    }
    public class IndexDifferentialFormat
    {
        public IndexDifferentialFormat()
        {
        }

        public UInt32 Index { get; set; }
        public DifferentialFormat DifferentialFormat { get; set; }
    }
    public class IndexCellStyle
    {
        public IndexCellStyle()
        {
        }

        public UInt32 Index { get; set; }
        public string Name { get; set; }
        public UInt32 FormatId { get; set; }
        public UInt32 BuildinId { get; set; }
        public CellStyle CellStyle { get; set; }
    }
    public class IndexFormat
    {
        public IndexFormat()
        {
        }

        public UInt32 Index { get; set; }
        public Format Format { get; set; }
    }
    public class SheetNameAndID
    {
        public string SheetName { get; set; }
        public string SheetID { get; set; }
    }

    // Enum Excel
    public enum FormulaEnum
    {
        Error = 0,
        NONE = 1,
        SUM = 2,
        AVERAGE = 3,
    }
    public enum FontNameEnum
    {
        Error = 0,
        Arial = 1,
        CourrierNew = 2,
    }
}
