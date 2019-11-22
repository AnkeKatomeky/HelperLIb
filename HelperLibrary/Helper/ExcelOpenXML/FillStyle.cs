using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary.ExcelOpenXml
{
    public class StyleFill
    {
        public UInt32Value FillIndex { get; set; }
        public FillColor Color { get; set; }

        public StyleFill(FillColor color)
        {
            Color = color;
        }

        public void SetFill(WorkbookStylesPart stylesPart, UInt32Value fillIndex)
        {
            var fill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill()
            {
                PatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid
            };

            fill.ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor
            {
                Rgb = HexBinaryValue.FromString(((uint)Color).ToString("x8"))
            };

            fill.BackgroundColor = new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor
            {
                Indexed = 64
            };

            stylesPart.Stylesheet.Fills.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Fill
            {
                PatternFill = fill
            });
            FillIndex = fillIndex;
        }
    }
}
