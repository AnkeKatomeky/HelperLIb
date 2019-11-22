using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary.ExcelOpenXml
{
    public class StyleExcel
    {
        public UInt32Value StyleIndex { get; set; }
        public StyleFont Font { get; set; }
        public StyleFill Fill { get; set; }
        public CellFormat Format { get; set; }
        public VerticalAlignmentValues Vertical { get; set; }
        public HorizontalAlignmentValues Horizontal { get; set; }
        public bool IsWordWrap { get; set; }
        public bool IsBorder { get; set; }

        public StyleExcel(StyleFont font, StyleFill fill, CellFormat format, VerticalAlignmentValues vertical, HorizontalAlignmentValues horizontal, bool iswordWrap)
        {
            Font = font;
            Fill = fill;
            Format = format;
            Vertical = vertical;
            Horizontal = horizontal;
            IsWordWrap = iswordWrap;
            IsBorder = false;
        }
        public StyleExcel(StyleFont font, StyleFill fill, CellFormat format, VerticalAlignmentValues vertical, HorizontalAlignmentValues horizontal, bool iswordWrap, bool isBorder)
        {
            Font = font;
            Fill = fill;
            Format = format;
            Vertical = vertical;
            Horizontal = horizontal;
            IsWordWrap = iswordWrap;
            IsBorder = isBorder;
        }

        public void SetStyle(WorkbookStylesPart stylesPart, UInt32Value index)
        {
            DocumentFormat.OpenXml.Spreadsheet.CellFormat cellFormat;
            NumberingFormat numberingFormat;
            Alignment aligment;

            cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();

            #region FormatsSwitch
            switch (Format)
            {
                case CellFormat.Header:
                    break;
                case CellFormat.String:
                    break;
                case CellFormat.StringWrap:
                    break;
                case CellFormat.NumberInt:
                    cellFormat.NumberFormatId = 3; // #,##0
                    cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
                    break;
                case CellFormat.NumberFract:
                    cellFormat.NumberFormatId = 4; // #,##0.00
                    cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
                    break;
                case CellFormat.NumberFractLong:
                    numberingFormat = new NumberingFormat();
                    numberingFormat.NumberFormatId = 164;
                    numberingFormat.FormatCode = StringValue.FromString("#0.000");

                    stylesPart.Stylesheet.NumberingFormats.AppendChild(numberingFormat);

                    cellFormat.NumberFormatId = numberingFormat.NumberFormatId; // 22; // m/d/yy h:mm
                    cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);


                    //cellFormat.NumberFormatId = 4; // #,##0.00   4
                    //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
                    break;

                case CellFormat.PercentInt:
                    cellFormat.NumberFormatId = 9; // 0%
                    cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
                    break;
                case CellFormat.PercentFract:
                    cellFormat.NumberFormatId = 10; // 0.00%
                    cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
                    break;
                case CellFormat.Date:
                    numberingFormat = new NumberingFormat();
                    numberingFormat.NumberFormatId = 165;
                    numberingFormat.FormatCode = StringValue.FromString("dd.mm.yyyy");

                    stylesPart.Stylesheet.NumberingFormats.AppendChild(numberingFormat);

                    cellFormat.NumberFormatId = numberingFormat.NumberFormatId; // 22; // m/d/yy h:mm
                    cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
                    break;
                case CellFormat.DateTime:
                    numberingFormat = new NumberingFormat();
                    numberingFormat.NumberFormatId = 166;
                    numberingFormat.FormatCode = StringValue.FromString("dd.mm.yyyy hh:mm");

                    stylesPart.Stylesheet.NumberingFormats.AppendChild(numberingFormat);

                    cellFormat.NumberFormatId = numberingFormat.NumberFormatId; // 22; // m/d/yy h:mm
                    cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
                    break;
                case CellFormat.Bool:
                    break;
                default:
                    break;
            }
            #endregion

            //if (!IsLock)
            //{
            //    cellFormat.Protection = new Protection() { Locked = false };
            //    cellFormat.ApplyProtection = true;

            //    // = new DocumentFormat.OpenXml.Spreadsheet.CellFormat() { ApplyProtection = true, Protection = new Protection() { Locked = false } };
            //}

            if (IsBorder)
            {
                cellFormat.BorderId = 1;
                cellFormat.ApplyBorder = true;
            }

            if (Font != null)
            {
                cellFormat.FontId = Font.FontIndex;
            }
            else
            {
                cellFormat.FontId = 0;
            }
            cellFormat.ApplyFont = true;

            if (Fill != null)
            {
                cellFormat.FillId = Fill.FillIndex;
            }
            else
            {
                cellFormat.FillId = 0;
            }
            cellFormat.ApplyFill = true;

            aligment = new Alignment();

            aligment.Vertical = Vertical;

            aligment.Horizontal = Horizontal;

            aligment.WrapText = IsWordWrap;
            cellFormat.AppendChild(aligment);

            stylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            StyleIndex = index;
        }
    }
}
