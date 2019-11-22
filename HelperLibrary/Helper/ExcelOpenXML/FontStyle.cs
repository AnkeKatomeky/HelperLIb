using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary.ExcelOpenXml
{
    public class StyleFont
    {
        public UInt32Value FontIndex { get; set; }
        public FontColor Color { get; set; }
        public bool IsBold { get; set; }

        public StyleFont(FontColor color, bool isBold)
        {
            Color = color;
            IsBold = isBold;
        }

        public void SetFont(WorkbookStylesPart stylesPart, UInt32Value fontIndex)
        {
            DocumentFormat.OpenXml.Spreadsheet.Font font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            if (IsBold)
            {
                font.Bold = new DocumentFormat.OpenXml.Spreadsheet.Bold();
            }
            font.Color = new DocumentFormat.OpenXml.Spreadsheet.Color()
            {
                Rgb = HexBinaryValue.FromString(((uint)Color).ToString("x8"))
            };           

            stylesPart.Stylesheet.Fonts.AppendChild(font);
            FontIndex = fontIndex;
        }

    }
}
