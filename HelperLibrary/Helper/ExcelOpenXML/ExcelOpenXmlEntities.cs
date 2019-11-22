using HelperLibrary;
using ReportGenerator.Model;
using System;
using System.Collections.Generic;

namespace HelperLibrary.ExcelOpenXml
{
    /// <summary>
    /// Excel spreadsheet column.
    /// </summary>
    public class ExcelColumn
    {
        /// <summary>
        /// Column constructor.
        /// </summary>
        /// <param name="title">Title text.</param>
        /// <param name="width">Width.</param>
        public ExcelColumn(string title, int width, CellColor style)
        {
            Title = title;
            Width = width;
            Style = style;
        }

        /// <summary>
        /// Column constructor.
        /// </summary>
        /// <param name="title">Title text.</param>
        /// <param name="width">Width.</param>
        public ExcelColumn(string title, int width, CellColor style, string group)
        {
            Title = title;
            Width = width;
            Style = style;
            Group = group;
        }

        /// <summary>
        /// Title text.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Width.
        /// </summary>
        public int Width { get; set; }

        public CellColor Style { get; set; }
        public string Group { get; set; }
    }

    /// <summary>
    /// Abstract excel spreadsheet cell.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Too many classes.")]
    public abstract class ExcelCell
    {
        private List<IfFormat> useErrorStyle;
        private string formula;
        public SummuryUBD Reference { get; set; }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCell(List<IfFormat> useErrorStyle)
        {
            this.useErrorStyle = useErrorStyle;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCell(List<IfFormat> useErrorStyle, SummuryUBD refe)
        {
            this.useErrorStyle = useErrorStyle;
            Reference = refe;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCell(List<IfFormat> useErrorStyle, string formula)
        {
            this.useErrorStyle = useErrorStyle;
            this.formula = formula;
        }

        /// <summary>
        /// Cell should display an error.
        /// </summary>
        public List<IfFormat> UseErrorStyle
        {
            get
            {
                return useErrorStyle;
            }
        }
        public string Formula
        { 
            get
            {
                return formula;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with string value.
    /// </summary>
    public class ExcelCellString : ExcelCell
    {
        private string stringValue;
        private bool wordWrap;
        private List<string> dataValidation;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">String value.</param>
        public ExcelCellString(string value)
            : base(null)
        {
            stringValue = value;
        }
        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">String value.</param>
        public ExcelCellString(string value,string formula)
            : base(null, formula)
        {
            stringValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">String value.</param>
        public ExcelCellString(string value, List<string> dataValidation)
            : base(null)
        {
            this.dataValidation = dataValidation;
            stringValue = value;
        }


        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">String value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        /// <param name="wordWrap">Cell text should be wrapped.</param>
        public ExcelCellString(string value, List<IfFormat> useErrorStyle, SummuryUBD refe)
            : base(useErrorStyle, refe)
        {
            stringValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">String value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        /// <param name="wordWrap">Cell text should be wrapped.</param>
        public ExcelCellString(string value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            stringValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">String value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        /// <param name="wordWrap">Cell text should be wrapped.</param>
        public ExcelCellString(string value, List<IfFormat> useErrorStyle,bool wordWrap)
            : base(useErrorStyle)
        {
            stringValue = value;
            this.wordWrap = wordWrap;
        }


        /// <summary>
        /// Cell value.
        /// </summary>
        public string Value
        {
            get
            {
                return stringValue;
            }
        }

        public List<string> DataValidation
        {
            get
            {
                return dataValidation;
            }
        }

        /// <summary>
        /// Cell text should be wrapped.
        /// </summary>
        public bool WordWrap
        {
            get
            {
                return wordWrap;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with boolean value.
    /// </summary>
    public class ExcelCellBoolean : ExcelCell
    {
        private bool? booleanValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Boolean value.</param>
        public ExcelCellBoolean(bool? value)
            : base(null)
        {
            booleanValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Boolean value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellBoolean(bool? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            booleanValue = value;
        }

        /// <summary>
        /// Cell value.
        /// </summary>
        public bool? Value
        {
            get
            {
                return booleanValue;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with number value (integral).
    /// </summary>        
    public class ExcelCellNumberIntegral : ExcelCell
    {
        private int? numberIntegralValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Integer value.</param>
        public ExcelCellNumberIntegral(int? value)
            : base(null)
        {
            numberIntegralValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Integer value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellNumberIntegral(int? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            numberIntegralValue = value;
        }

                /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellNumberIntegral(int? value, string formula)
            : base(null,formula)
        {
            numberIntegralValue = value;
        }


        /// <summary>
        /// Cell value.
        /// </summary>
        public int? Value
        {
            get
            {
                return numberIntegralValue;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with number value (fractional).
    /// </summary>        
    public class ExcelCellNumberFractional : ExcelCell
    {
        private decimal? numberFractionalValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellNumberFractional(decimal? value)
            : base(null)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellNumberFractional(decimal? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellNumberFractional(decimal? value, List<IfFormat> useErrorStyle, SummuryUBD refe)
            : base(useErrorStyle, refe)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellNumberFractional(decimal? value, string formula)
            : base(null,formula)
        {
            numberFractionalValue = value;
        }


        /// <summary>
        /// Cell value.
        /// </summary>
        public decimal? Value
        {
            get
            {
                return numberFractionalValue;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with number value (fractional).
    /// </summary>        
    public class ExcelCellNumberFractionalHuge : ExcelCell
    {
        private double? numberFractionalValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellNumberFractionalHuge(double? value)
            : base(null)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellNumberFractionalHuge(double? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellNumberFractionalHuge(double? value, List<IfFormat> useErrorStyle, SummuryUBD refe)
            : base(useErrorStyle, refe)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellNumberFractionalHuge(double? value, string formula)
            : base(null, formula)
        {
            numberFractionalValue = value;
        }


        /// <summary>
        /// Cell value.
        /// </summary>
        public double? Value
        {
            get
            {
                return numberFractionalValue;
            }
        }
    }


    /// <summary>
    /// Excel spreadsheet cell with number value (fractional).
    /// </summary>        
    public class ExcelCellNumberFractionalLong : ExcelCell
    {
        private decimal? numberFractionalValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellNumberFractionalLong(decimal? value)
            : base(null)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellNumberFractionalLong(decimal? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellNumberFractionalLong(decimal? value, List<IfFormat> useErrorStyle, SummuryUBD refe)
            : base(useErrorStyle, refe)
        {
            numberFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellNumberFractionalLong(decimal? value, string formula)
            : base(null, formula)
        {
            numberFractionalValue = value;
        }


        /// <summary>
        /// Cell value.
        /// </summary>
        public decimal? Value
        {
            get
            {
                return numberFractionalValue;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with percent value (integral).
    /// </summary>        
    public class ExcelCellPercentIntegral : ExcelCell
    {
        private decimal? percentIntegralValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellPercentIntegral(decimal? value)
            : base(null)
        {
            percentIntegralValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellPercentIntegral(decimal? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            percentIntegralValue = value;
        }

        /// <summary>
        /// Cell value.
        /// </summary>
        public decimal? Value
        {
            get
            {
                return percentIntegralValue;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with percent value (fractional).
    /// </summary>        
    public class ExcelCellPercentFractional : ExcelCell
    {
        private decimal? percentFractionalValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        public ExcelCellPercentFractional(decimal? value)
            : base(null)
        {
            percentFractionalValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Decimal value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellPercentFractional(decimal? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            percentFractionalValue = value;
        }

        /// <summary>
        /// Cell value.
        /// </summary>
        public decimal? Value
        {
            get
            {
                return percentFractionalValue;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with date value.
    /// </summary>        
    public class ExcelCellDate : ExcelCell
    {
        private DateTime? dateValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Date value.</param>
        public ExcelCellDate(DateTime? value)
            : base(null)
        {
            dateValue = value;
        }
        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Date value.</param>
        public ExcelCellDate(DateTime? value,string formula)
            : base(null, formula)
        {
            dateValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Date value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellDate(DateTime? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            dateValue = value;
        }

        /// <summary>
        /// Cell value.
        /// </summary>
        public DateTime? Value
        {
            get
            {
                return dateValue;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet cell with date and time value.
    /// </summary>        
    public class ExcelCellDateAndTime : ExcelCell
    {
        private DateTime? dateAndTimeValue;

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Date and time value.</param>
        public ExcelCellDateAndTime(DateTime? value)
            : base(null)
        {
            dateAndTimeValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Date and time value.</param>
        public ExcelCellDateAndTime(DateTime? value,string formula)
            : base(null, formula)
        {
            dateAndTimeValue = value;
        }

        /// <summary>
        /// Cell constructor.
        /// </summary>
        /// <param name="value">Date and time value.</param>
        /// <param name="useErrorStyle">Cell should display an error.</param>
        public ExcelCellDateAndTime(DateTime? value, List<IfFormat> useErrorStyle)
            : base(useErrorStyle)
        {
            dateAndTimeValue = value;
        }

        /// <summary>
        /// Cell value.
        /// </summary>
        public DateTime? Value
        {
            get
            {
                return dateAndTimeValue;
            }
        }
    }

    /// <summary>
    /// Excel spreadsheet column definition.
    /// </summary>
    public class ExcelColumnDefinition
    {
        private string caption;
        private ExcelCell cellValue;
        private int index;

        public bool IsManyVariant { get; set; }
        public int Skip { get; set; }

        /// <summary>
        /// Variants of capations 
        /// </summary>
        public string[] Variants
        {
            get
            {
                if (IsManyVariant)
                {
                    return caption.Split('|');
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Column definition constructor.
        /// </summary>
        /// <param name="caption">Column caption.</param>
        /// <param name="cellValue">Cell typed value.</param>
        public ExcelColumnDefinition(string caption, ExcelCell cellValue)
        {
            this.caption = caption;
            this.cellValue = cellValue;
            IsManyVariant = false;
            index = 0;
        }

        /// <summary>
        /// Column definition constructor whith seted more than one capation(capation separeted by "|" cart)
        /// </summary>
        /// <param name="caption">Column caption.</param>
        /// <param name="cellValue">Cell typed value.</param>
        /// <param name="isManyVariant">Is capation seted as many variants(capation separeted by "|" cart)</param>
        public ExcelColumnDefinition(string caption, ExcelCell cellValue, bool isManyVariant)
        {
            this.caption = caption;
            this.cellValue = cellValue;
            IsManyVariant = isManyVariant;
            index = 0;
        }

        /// <summary>
        /// Column definition constructor if more than one same capation in a row
        /// </summary>
        /// <param name="caption">Column caption.</param>
        /// <param name="cellValue">Cell typed value.</param>
        /// <param name="skip">How many cells whith the same capation need to skip</param>
        public ExcelColumnDefinition(string caption, ExcelCell cellValue, int skip)
        {
            this.caption = caption;
            this.cellValue = cellValue;
            Skip = skip;
            IsManyVariant = false;
            index = 0;
        }


        /// <summary>
        /// Column definition constructor if more than one same capation in a row and whith seted more than one capation(capation separeted by "|" cart)
        /// </summary>
        /// <param name="caption">Column caption.</param>
        /// <param name="cellValue">Cell typed value.</param>
        /// <param name="skip">How many cells whith the same capation need to skip</param>
        /// <param name="isManyVariant">Is capation seted as many variants(capation separeted by "|" cart)</param>
        public ExcelColumnDefinition(string caption, ExcelCell cellValue, int skip, bool isManyVariant)
        {
            this.caption = caption;
            this.cellValue = cellValue;
            Skip = skip;
            IsManyVariant = isManyVariant;
            index = 0;
        }

        /// <summary>
        /// Column caption.
        /// </summary>
        public string Caption
        {
            get
            {
                return caption;
            }
        }

        /// <summary>
        /// Column index in worksheet.
        /// </summary>
        public int Index
        {
            get
            {
                return index;
            }

            set
            {
                index = value;
            }
        }

        /// <summary>
        /// Cell typed value.
        /// </summary>
        public ExcelCell CellValue
        {
            get
            {
                return cellValue;
            }

            set
            {
                cellValue = value;
            }
        }
    }
}
