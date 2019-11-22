using System.Collections.Generic;
using DocumentFormat.OpenXml;

namespace HelperLibrary.ExcelOpenXml
{
    public enum CellColor
    {
        Default,
        Purple,
        PaleGreen,
        PaleGreenAndRed,
        Green,
        WGreen,
        Yellow,
        Red,
        Blue,
        DBlue,
        Marsh,
        Brown,
        Pastel,
        LGreen,
        Pink,
        LBrown,
        Cyian,
        DefaultB,
        GreenB,
        YellowB,
        RedB,
        BlueB,
        PaleRed,
        BlueAndRed,
        WBlue,
        GrayB,
        Gray,
        DGray
    }

    public enum FillColor : uint
    {
        Default = 0xffd9d9d9,
        Gray = 0xffa6a6a6,
        Purple = 0xffccc0da,
        Green = 0xffc6efce,
        WGreen = 0xffd8e4bc,
        Yellow = 0xffffeb9c,
        Red = 0xffffc7ce,
        Blue = 0xffc8d8f0,
        DBlue = 0xff4f81bd,
        Marsh = 0xffa9d79b,
        Brown = 0xffda9694,
        Pastel = 0xfffde9d9,
        LGreen = 0xffd8e4bc,
        Pink = 0xfff2dcdb,
        LBrown = 0xfffcd5b4,
        Cyian = 0xff317c9b,
        PaleGreen = 0xffedfaf0,
        PaleRed = 0xffffdbdf,
        WBlue = 0xffdfe8f2
    }

    public enum FontColor : uint
    {
        Default = 0xff282828,
        Gray = 0xff0f0f0f,        
        Purple = 0xff342247,
        Green = 0xff006100,
        Yellow = 0xff9c6500,
        Red = 0xff9c0006,
        RedFull = 0xffff0000,
        Blue = 0xff202020
    }

    public enum CellFormat
    {
        Header,
        String, StringWrap,
        NumberInt, NumberFract, NumberFractLong,
        PercentInt, PercentFract,
        Date, DateTime,
        Bool
    }



    public class ExcelOpenXmlSaxWriterOptions
    {
        public static ExcelOpenXmlSaxWriterOptions Instanse
        {
            get
            {
                if (_inst == null)
                {
                    _inst = new ExcelOpenXmlSaxWriterOptions();
                }
                return _inst;
            }
        }
        private static ExcelOpenXmlSaxWriterOptions _inst;

        public List<StyleFont> Fonts { get; set; }
        public List<StyleFill> Fills { get; set; }
        public List<StyleExcel> Styles { get; set; }

        #region FontProp
        private StyleFont DefaultFont { get; set; }
        private StyleFont PurpleFont { get; set; }
        private StyleFont GreenFont { get; set; }
        private StyleFont YellowFont { get; set; }
        private StyleFont RedFont { get; set; }
        private StyleFont RedFullFont { get; set; }
        private StyleFont BlueFont { get; set; }
        private StyleFont DBlueFont { get; set; }
        private StyleFont DGrayFont { get; set; }
        #endregion

        #region FillProp
        private StyleFill DefaultFill { get; set; }
        private StyleFill PurpleFill { get; set; }
        private StyleFill GreenFill { get; set; }
        private StyleFill WGreenFill { get; set; }
        private StyleFill YellowFill { get; set; }
        private StyleFill RedFill { get; set; }
        private StyleFill BlueFill { get; set; }
        private StyleFill DBlueFill { get; set; }
        private StyleFill MarshFill { get; set; }
        private StyleFill BrownFill { get; set; }
        private StyleFill PastelFill { get; set; }
        private StyleFill LGreenFill { get; set; }
        private StyleFill PinkFill { get; set; }
        private StyleFill LBrownFill { get; set; }
        private StyleFill CyianFill { get; set; }
        private StyleFill PaleGreenFill { get; set; }
        private StyleFill PaleRedFill { get; set; }
        private StyleFill WBlueFill { get; set; }
        private StyleFill DGrayFill { get; set; }
        #endregion

        #region HeaderProp
        public StyleExcel DefaultHeader { get; set; }
        public StyleExcel DefaultHeaderBorder { get; set; }
        public StyleExcel PurpleHeader { get; set; }
        public StyleExcel GreenHeader { get; set; }
        public StyleExcel WGreenHeader { get; set; }
        public StyleExcel YellowHeader { get; set; }
        public StyleExcel RedHeader { get; set; }
        public StyleExcel GreenHeaderBorder { get; set; }
        public StyleExcel YellowHeaderBorder { get; set; }
        public StyleExcel RedHeaderBorder { get; set; }
        public StyleExcel BlueHeader { get; set; }
        public StyleExcel BlueHeaderBorder { get; set; }
        public StyleExcel DBlueHeader { get; set; }
        public StyleExcel MarshHeader { get; set; }
        public StyleExcel BrownHeader { get; set; }
        public StyleExcel PastelHeader { get; set; }
        public StyleExcel LGreenHeader { get; set; }
        public StyleExcel PinkHeader { get; set; }
        public StyleExcel LBrownHeader { get; set; }
        public StyleExcel CyianHeader { get; set; }
        public StyleExcel PaleGreenHeader { get; set; }
        public StyleExcel PaleGreenAndRedHeader { get; set; }
        public StyleExcel PaleRedHeader { get; set; }
        public StyleExcel BlueAndRedHeader { get; set; }
        #endregion

        #region IntProp
        public StyleExcel DefaultInt { get; set; }
        public StyleExcel DefaultIntBorder { get; set; }
        public StyleExcel BlueIntBorder { get; set; }
        public StyleExcel PurpleInt { get; set; }
        public StyleExcel GreenInt { get; set; }
        public StyleExcel YellowInt { get; set; }
        public StyleExcel RedInt { get; set; }
        public StyleExcel BlueInt { get; set; }
        public StyleExcel GrayIntBorder { get; set; }
        public StyleExcel GrayInt { get; set; }
        public StyleExcel PaleGreenInt { get; set; }
        public StyleExcel PaleRedInt { get; set; }
        public StyleExcel WBlueInt { get; set; }
        public StyleExcel DGrayInt { get; set; }
        public StyleExcel DBlueInt { get; set; }
        #endregion

        #region FractProp
        public StyleExcel DefaultFract { get; set; }
        public StyleExcel DefaultFractLong { get; set; }
        public StyleExcel DefaultFractBorder { get; set; }
        public StyleExcel PurpleFract { get; set; }
        public StyleExcel GreenFract { get; set; }
        public StyleExcel YellowFract { get; set; }
        public StyleExcel RedFract { get; set; }
        public StyleExcel GreenFractBorder { get; set; }
        public StyleExcel YellowFractBorder { get; set; }
        public StyleExcel RedFractBorder { get; set; }
        public StyleExcel BlueFract { get; set; }
        public StyleExcel BlueFractBorder { get; set; }
        public StyleExcel GrayFractBorder { get; set; }
        public StyleExcel GrayFract { get; set; }
        public StyleExcel PaleGreenFract { get; set; }
        public StyleExcel PaleRedFract { get; set; }
        public StyleExcel DGrayFract { get; set; }
        public StyleExcel DBlueFract { get; set; }
        #endregion

        #region PercentIntProp
        public StyleExcel PercentDefaultInt { get; set; }
        public StyleExcel PercentPurpleInt { get; set; }
        public StyleExcel PercentGreenInt { get; set; }
        public StyleExcel PercentYellowInt { get; set; }
        public StyleExcel PercentRedInt { get; set; }
        public StyleExcel PercentBlueInt { get; set; }
        #endregion

        #region PercentFractProp
        public StyleExcel PercentDefaultFract { get; set; }
        public StyleExcel PercentPurpleFract { get; set; }
        public StyleExcel PercentGreenFract { get; set; }
        public StyleExcel PercentYellowFract { get; set; }
        public StyleExcel PercentRedFract { get; set; }
        public StyleExcel PercentBlueFract { get; set; }
        #endregion

        #region Date
        public StyleExcel Date { get; set; }
        public StyleExcel PurpleDate { get; set; }
        public StyleExcel GreenDate { get; set; }
        public StyleExcel YellowDate { get; set; }
        public StyleExcel RedDate { get; set; }
        public StyleExcel BlueDate { get; set; }
        #endregion

        #region ETCProp
        public StyleExcel Boolean { get; set; }
        public StyleExcel WrapString { get; set; }
        public StyleExcel DateAndTime { get; set; }
        public StyleExcel DefaultString { get; set; }
        public StyleExcel DefaultStringBorder { get; set; }
        public StyleExcel PurpleString { get; set; }
        public StyleExcel GreenString { get; set; }
        public StyleExcel YellowString { get; set; }
        public StyleExcel RedString { get; set; }
        public StyleExcel BlueString { get; set; }
        public StyleExcel BlueStringBorder { get; set; }
        public StyleExcel MarshString { get; set; }
        public StyleExcel BrownString { get; set; }
        public StyleExcel PastelString { get; set; }
        public StyleExcel LGreenString { get; set; }
        public StyleExcel PinkString { get; set; }
        public StyleExcel LBrownString { get; set; }
        public StyleExcel GrayStringBorder { get; set; }
        public StyleExcel PaleGreenString { get; set; }
        public StyleExcel PaleRedString { get; set; }
        public StyleExcel WBlueString { get; set; }
        public StyleExcel DGrayString { get; set; }
        public StyleExcel DBlueString { get; set; }
        #endregion

        private ExcelOpenXmlSaxWriterOptions()
        {
            #region FontEmpli
            DefaultFont = new StyleFont(FontColor.Default, true);
            PurpleFont = new StyleFont(FontColor.Purple, true);
            GreenFont = new StyleFont(FontColor.Green, true);
            YellowFont = new StyleFont(FontColor.Yellow, true);
            RedFont = new StyleFont(FontColor.Red, true);
            RedFullFont = new StyleFont(FontColor.RedFull, true);
            BlueFont = new StyleFont(FontColor.Blue, true);
            DGrayFont = new StyleFont(FontColor.Gray, true);
            #endregion
            #region FillEmpli
            DefaultFill = new StyleFill(FillColor.Default);
            PurpleFill = new StyleFill(FillColor.Purple);
            GreenFill = new StyleFill(FillColor.Green);
            PaleGreenFill = new StyleFill(FillColor.PaleGreen);
            WGreenFill = new StyleFill(FillColor.WGreen);
            YellowFill = new StyleFill(FillColor.Yellow);
            RedFill = new StyleFill(FillColor.Red);
            PaleRedFill = new StyleFill(FillColor.PaleRed);
            BlueFill = new StyleFill(FillColor.Blue);
            WBlueFill = new StyleFill(FillColor.WBlue);
            DBlueFill = new StyleFill(FillColor.DBlue);
            MarshFill = new StyleFill(FillColor.Marsh);
            BrownFill = new StyleFill(FillColor.Brown);
            PastelFill = new StyleFill(FillColor.Pastel);
            LGreenFill = new StyleFill(FillColor.LGreen);
            PinkFill = new StyleFill(FillColor.Pink);
            LBrownFill = new StyleFill(FillColor.LBrown);
            CyianFill = new StyleFill(FillColor.Cyian);
            DGrayFill = new StyleFill(FillColor.Gray);
            #endregion
            #region HeaderEmpli
            DefaultHeaderBorder = new StyleExcel(DefaultFont, DefaultFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true, true);
            GreenHeaderBorder = new StyleExcel(GreenFont, GreenFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true, true);
            YellowHeaderBorder = new StyleExcel(YellowFont, YellowFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true, true);
            RedHeaderBorder = new StyleExcel(RedFont, RedFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true, true);
            BlueHeaderBorder = new StyleExcel(BlueFont, BlueFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true, true);

            DefaultHeader = new StyleExcel(DefaultFont, DefaultFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            PaleGreenHeader = new StyleExcel(DefaultFont, PaleGreenFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            PaleGreenAndRedHeader = new StyleExcel(RedFullFont, PaleGreenFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            PurpleHeader = new StyleExcel(PurpleFont, PurpleFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            GreenHeader = new StyleExcel(GreenFont, GreenFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            WGreenHeader = new StyleExcel(GreenFont, WGreenFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            YellowHeader = new StyleExcel(YellowFont, YellowFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            RedHeader = new StyleExcel(RedFont, RedFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            PaleRedHeader = new StyleExcel(RedFullFont, PaleRedFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            BlueHeader = new StyleExcel(BlueFont, BlueFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            BlueAndRedHeader = new StyleExcel(RedFullFont, BlueFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            DBlueHeader = new StyleExcel(BlueFont, DBlueFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            MarshHeader = new StyleExcel(DefaultFont, MarshFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            BrownHeader = new StyleExcel(DefaultFont, BrownFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            PastelHeader = new StyleExcel(DefaultFont, PastelFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            LGreenHeader = new StyleExcel(DefaultFont, LGreenFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            PinkHeader = new StyleExcel(DefaultFont, PinkFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            LBrownHeader = new StyleExcel(DefaultFont, LBrownFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            CyianHeader = new StyleExcel(DefaultFont, CyianFill, CellFormat.Header, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, true);
            #endregion
            #region IntEmpli
            DefaultInt = new StyleExcel(null, null, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            DefaultIntBorder = new StyleExcel(null, null, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            GrayIntBorder = new StyleExcel(DefaultFont, DefaultFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            GrayInt = new StyleExcel(DefaultFont, DefaultFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, false);
            DGrayInt = new StyleExcel(DGrayFont, DGrayFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, false);
            DBlueInt = new StyleExcel(BlueFont, DBlueFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, false);

            PurpleInt = new StyleExcel(PurpleFont, PurpleFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            GreenInt = new StyleExcel(GreenFont, GreenFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PaleGreenInt = new StyleExcel(GreenFont, PaleGreenFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            YellowInt = new StyleExcel(YellowFont, YellowFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            RedInt = new StyleExcel(RedFont, RedFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PaleRedInt = new StyleExcel(RedFont, PaleRedFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            BlueInt = new StyleExcel(BlueFont, BlueFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            WBlueInt = new StyleExcel(BlueFont, WBlueFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            BlueIntBorder = new StyleExcel(BlueFont, BlueFill, CellFormat.NumberInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            #endregion
            #region FractEmpli
            DefaultFract = new StyleExcel(null, null, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            DefaultFractLong = new StyleExcel(null, null, CellFormat.NumberFractLong, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PurpleFract = new StyleExcel(PurpleFont, PurpleFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            GreenFract = new StyleExcel(GreenFont, GreenFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PaleGreenFract = new StyleExcel(GreenFont, PaleGreenFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            YellowFract = new StyleExcel(YellowFont, YellowFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            RedFract = new StyleExcel(RedFont, RedFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PaleRedFract = new StyleExcel(RedFont, PaleRedFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            BlueFract = new StyleExcel(BlueFont, BlueFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            DGrayFract = new StyleExcel(DGrayFont, DGrayFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            DBlueFract = new StyleExcel(BlueFont, DBlueFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);

            BlueFractBorder = new StyleExcel(BlueFont, BlueFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            DefaultFractBorder = new StyleExcel(null, null, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            GreenFractBorder = new StyleExcel(GreenFont, GreenFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            GrayFractBorder = new StyleExcel(DefaultFont, DefaultFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            GrayFract = new StyleExcel(DefaultFont, DefaultFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, false);
            GreenFractBorder = new StyleExcel(GreenFont, GreenFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            YellowFractBorder = new StyleExcel(YellowFont, YellowFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            RedFractBorder = new StyleExcel(RedFont, RedFill, CellFormat.NumberFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            #endregion
            #region PercentIntEmpli
            PercentDefaultInt = new StyleExcel(null, null, CellFormat.PercentInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentPurpleInt = new StyleExcel(PurpleFont, PurpleFill, CellFormat.PercentInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentGreenInt = new StyleExcel(GreenFont, GreenFill, CellFormat.PercentInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentYellowInt = new StyleExcel(YellowFont, YellowFill, CellFormat.PercentInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentRedInt = new StyleExcel(RedFont, RedFill, CellFormat.PercentInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentBlueInt = new StyleExcel(BlueFont, BlueFill, CellFormat.PercentInt, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            #endregion
            #region PercentFractEmpli
            PercentDefaultFract = new StyleExcel(null, null, CellFormat.PercentFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentPurpleFract = new StyleExcel(PurpleFont, PurpleFill, CellFormat.PercentFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentGreenFract = new StyleExcel(GreenFont, GreenFill, CellFormat.PercentFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentYellowFract = new StyleExcel(YellowFont, YellowFill, CellFormat.PercentFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentRedFract = new StyleExcel(RedFont, RedFill, CellFormat.PercentFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PercentBlueFract = new StyleExcel(BlueFont, BlueFill, CellFormat.PercentFract, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            #endregion
            #region ETCEmpli
            Boolean = new StyleExcel(null, null, CellFormat.Bool, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);
            WrapString = new StyleExcel(null, null, CellFormat.StringWrap, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, true);
            Date = new StyleExcel(null, null, CellFormat.Date, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);
            DateAndTime = new StyleExcel(null, null, CellFormat.DateTime, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);

            DefaultString = new StyleExcel(null, null, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            DefaultStringBorder = new StyleExcel(null, null, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            GrayStringBorder = new StyleExcel(DefaultFont, DefaultFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            PurpleString = new StyleExcel(null, PurpleFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            GreenString = new StyleExcel(null, GreenFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PaleGreenString = new StyleExcel(null, PaleGreenFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            YellowString = new StyleExcel(null, YellowFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            RedString = new StyleExcel(null, RedFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PaleRedString = new StyleExcel(null, PaleRedFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            BlueString = new StyleExcel(null, BlueFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            WBlueString = new StyleExcel(null, WBlueFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            BlueStringBorder = new StyleExcel(null, BlueFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            DGrayString = new StyleExcel(null, DGrayFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            DBlueString = new StyleExcel(null, DBlueFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false, true);
            MarshString = new StyleExcel(null, MarshFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            BrownString = new StyleExcel(null, BrownFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PastelString = new StyleExcel(null, PastelFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            LGreenString = new StyleExcel(null, LGreenFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            PinkString = new StyleExcel(null, PinkFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            LBrownString = new StyleExcel(null, LBrownFill, CellFormat.String, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General, false);
            #endregion
            #region Date
            Date = new StyleExcel(null, null, CellFormat.Date, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);
            PurpleDate = new StyleExcel(null, PurpleFill, CellFormat.Date, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);
            GreenDate = new StyleExcel(null, GreenFill, CellFormat.Date, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);
            YellowDate = new StyleExcel(null, YellowFill, CellFormat.Date, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);
            RedDate = new StyleExcel(null, RedFill, CellFormat.Date, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);
            BlueDate = new StyleExcel(null, BlueFill, CellFormat.Date, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center, false);
            #endregion
            Fonts = new List<StyleFont>()
            {
                DefaultFont,
                PurpleFont,
                GreenFont,
                YellowFont,
                RedFont,
                RedFullFont,
                BlueFont,
                DGrayFont,
            };
            Fills = new List<StyleFill>()
            {
                DefaultFill,
                PurpleFill,
                GreenFill,
                PaleGreenFill,
                WGreenFill,
                YellowFill,
                RedFill,
                PaleRedFill,
                BlueFill,
                WBlueFill,
                DBlueFill,
                MarshFill,
                BrownFill,
                PastelFill,
                LGreenFill,
                PinkFill,
                LBrownFill,
                CyianFill,
                DGrayFill,
            };
            Styles = new List<StyleExcel>()
            {
                DefaultString, PurpleString, GreenString,
                PaleGreenString, YellowString, RedString,
                BlueString, BlueStringBorder, MarshString,
                BrownString, PastelString, LGreenString,
                PinkString, LBrownString, DefaultStringBorder,
                PaleRedString, WBlueString, GrayStringBorder,
                DGrayString, DBlueString,

                DefaultHeader, PurpleHeader, GreenHeader,
                YellowHeader, RedHeader, BlueHeader,
                DBlueHeader,  CyianHeader, MarshHeader,
                BrownHeader, PastelHeader, LGreenHeader,
                PinkHeader, LBrownHeader, DefaultHeaderBorder,
                GreenHeaderBorder, YellowHeaderBorder, BlueHeaderBorder,
                RedHeaderBorder, WGreenHeader, PaleRedHeader,
                BlueAndRedHeader, PaleGreenHeader, PaleGreenAndRedHeader,

                DefaultInt, DefaultIntBorder, PurpleInt,
                GreenInt, YellowInt, RedInt,
                BlueInt, BlueIntBorder, PaleGreenInt,
                PaleRedInt, GrayIntBorder, GrayInt, WBlueInt,
                DGrayInt, DBlueInt,

                DefaultFract, PurpleFract, GreenFract,
                YellowFract, RedFract, BlueFract,
                BlueFractBorder, DefaultFractBorder,
                GreenFractBorder, YellowFractBorder,
                RedFractBorder, DefaultFractLong,
                DefaultFractLong, PaleGreenFract, PaleRedFract,
                GrayFractBorder, GrayFract, DGrayFract,
                DBlueFract,

                PercentDefaultInt, PercentPurpleInt,
                PercentGreenInt, PercentYellowInt,
                PercentRedInt, PercentBlueInt,

                PercentDefaultFract, PercentPurpleFract,
                PercentGreenFract, PercentYellowFract,
                PercentRedFract, PercentBlueFract,

                Boolean, DefaultString, WrapString, DateAndTime,

                Date, PurpleDate, GreenDate, YellowDate, RedDate, BlueDate

            };


            FixTopRow = true;
            AddAutoFilter = true;
            FixLeftColums = false;
            LockColumn = "B2";
            Offset = 2d;
            HeaderHeight = 0;
        }

        public bool FixTopRow { get; set; }
        public bool FixLeftColums { get; set; }
        public string LockColumn { get; set; }
        public double Offset { get; set; }
        public bool AddAutoFilter { get; set; }
        public int HeaderHeight { get; set; }

        public bool IsToAddValidator { get; set; }
        public string ValidatorParams { get; set; }
        public string ValidatorColumn { get; set; }
    }
}
