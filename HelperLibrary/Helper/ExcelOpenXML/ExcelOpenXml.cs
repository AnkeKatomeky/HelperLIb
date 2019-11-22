using System;
using System.Collections.Generic;
using System.Linq;

namespace HelperLibrary.ExcelOpenXml
{
    /// <summary>
    /// Writes data to Excel workbooks via OpenXML (uses DOM, has file size limitation).
    /// </summary>
    public class ExcelOpenXml : IDisposable
    {
        private DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheetDocument;
        private DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData;

        private DocumentFormat.OpenXml.UInt32Value booleanFormatStyleIndex,
            numberIntegralFormatStyleIndex, numberFractionalFormatStyleIndex,
            percentIntegralFormatStyleIndex, percentFractionalFormatStyleIndex,
            dateFormatStyleIndex, dateTimeFormatStyleIndex;

        /// <summary>
        /// Frees associated resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }

        /// <summary>
        /// Opens workbook from file.
        /// </summary>
        public bool OpenWorkbook(string fileName)
        {
            CloseWorkbook();

            spreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(fileName, true);

            return spreadsheetDocument != null;
        }

        /// <summary>
        /// Closes workbook.
        /// </summary>
        public void CloseWorkbook()
        {
            Dispose();
        }

        /// <summary>
        /// Saves changes in workbook.
        /// </summary>
        public void SaveWorkbook()
        {
            if (spreadsheetDocument == null)
            {
                return;
            }

            ////this.workbook.WorkbookPart.WorksheetParts.First().Worksheet.Save();
            spreadsheetDocument.WorkbookPart.Workbook.Save();
        }

        /// <summary>
        /// Opens the first worksheet of current workbook.
        /// </summary>
        public bool OpenFirstWorksheet()
        {
            if (spreadsheetDocument == null)
            {
                return false;
            }

            sheetData = spreadsheetDocument.WorkbookPart.WorksheetParts.First().Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();

            return sheetData != null;
        }

        /// <summary>
        /// Opens specific worksheet of current workbook.
        /// </summary>
        public bool OpenWorksheet(int index)
        {
            if (spreadsheetDocument == null)
            {
                return false;
            }

            sheetData = null;

            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ElementAt(index);

            foreach (var worksheetPart in spreadsheetDocument.WorkbookPart.WorksheetParts)
            {
                if (spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart) == sheet.Id)
                {
                    sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                }
            }

            return sheetData != null;
        }

        /// <summary>
        /// Adds styles to workbook (for reporting).
        /// </summary>
        public void AddStylesForReport()
        {
            if (spreadsheetDocument == null)
            {
                return;
            }

            if (spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats == null)
            {
                spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats = new DocumentFormat.OpenXml.Spreadsheet.CellFormats();
            }

            if (spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats == null)
            {
                spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormats();
            }

            DocumentFormat.OpenXml.Spreadsheet.CellFormat cellFormat;
            DocumentFormat.OpenXml.Spreadsheet.NumberingFormat numberingFormat;

            DocumentFormat.OpenXml.UInt32Value customFormatId = 164;

            // Number (integral)
            cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            cellFormat.NumberFormatId = 3; // #,##0
            cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            numberIntegralFormatStyleIndex = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            // Number (fractional)
            cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            cellFormat.NumberFormatId = 4; // #,##0.00
            cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            numberFractionalFormatStyleIndex = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            // Percent (integral)
            cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            cellFormat.NumberFormatId = 9; // 0%
            cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            percentIntegralFormatStyleIndex = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            // Percent (fractional)
            cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            cellFormat.NumberFormatId = 10; // 0.00%
            cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            percentFractionalFormatStyleIndex = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            // Date
            numberingFormat = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat();
            numberingFormat.NumberFormatId = customFormatId++;
            numberingFormat.FormatCode = DocumentFormat.OpenXml.StringValue.FromString("dd.mm.yyyy");
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.AppendChild(numberingFormat);

            cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            cellFormat.NumberFormatId = numberingFormat.NumberFormatId; // 14; // mm-dd-yy
            cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment { Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center });
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            dateFormatStyleIndex = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            // Date and time
            numberingFormat = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat();
            numberingFormat.NumberFormatId = customFormatId++;
            numberingFormat.FormatCode = DocumentFormat.OpenXml.StringValue.FromString("dd.mm.yyyy hh:mm");
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.AppendChild(numberingFormat);

            cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            cellFormat.NumberFormatId = numberingFormat.NumberFormatId; // 22; // m/d/yy h:mm
            cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment { Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center });
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            dateTimeFormatStyleIndex = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            // Boolean            
            cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment { Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center });
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            booleanFormatStyleIndex = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
        }

        /// <summary>
        /// Opens workbook (for reporting).
        /// </summary>
        public bool OpenWorkbookForReport(string fileName)
        {
            if (!OpenWorkbook(fileName))
            {
                return false;
            }

            if (!OpenFirstWorksheet())
            {
                return false;
            }

            AddStylesForReport();

            return true;
        }

        /// <summary>
        /// Closes workbook (for reporting).
        /// </summary>
        public void CloseWorkbookForReport()
        {
            SaveWorkbook();
            CloseWorkbook();
        }

        /// <summary>
        /// Adds row of cells to current worksheet (for reporting).
        /// </summary>
        public void AddRowToReport(ExcelCell[] cells)
        {
            if (sheetData == null)
            {
                return;
            }

            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row();

            DocumentFormat.OpenXml.Spreadsheet.Cell cell;

            ExcelCellString stringCell;
            ExcelCellBoolean booleanCell;
            ExcelCellNumberIntegral numberIntegralCell;
            ExcelCellNumberFractional numberFractionalCell;
            ExcelCellPercentIntegral percentIntegralCell;
            ExcelCellPercentFractional percentFractionalCell;
            ExcelCellDate dateCell;
            ExcelCellDateAndTime dateAndTimeCell;

            foreach (ExcelCell excelCell in cells)
            {
                stringCell = excelCell as ExcelCellString;
                if (stringCell != null)
                {
                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(stringCell.Value);
                    row.AppendChild(cell);
                    continue;
                }

                booleanCell = excelCell as ExcelCellBoolean;
                if (booleanCell != null)
                {
                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.StyleIndex = booleanFormatStyleIndex;
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    if (booleanCell.Value.HasValue)
                    {
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(booleanCell.Value.Value ? "Да" : null);
                    }

                    row.AppendChild(cell);
                    continue;
                }

                numberIntegralCell = excelCell as ExcelCellNumberIntegral;
                if (numberIntegralCell != null)
                {
                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.StyleIndex = numberIntegralFormatStyleIndex;
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                    if (numberIntegralCell.Value.HasValue)
                    {
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(DocumentFormat.OpenXml.Int32Value.FromInt32(numberIntegralCell.Value.Value));
                    }

                    row.AppendChild(cell);
                    continue;
                }

                numberFractionalCell = excelCell as ExcelCellNumberFractional;
                if (numberFractionalCell != null)
                {
                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.StyleIndex = numberFractionalFormatStyleIndex;
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                    if (numberFractionalCell.Value.HasValue)
                    {
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(DocumentFormat.OpenXml.DecimalValue.FromDecimal(numberFractionalCell.Value.Value));
                    }

                    row.AppendChild(cell);
                    continue;
                }

                percentIntegralCell = excelCell as ExcelCellPercentIntegral;
                if (percentIntegralCell != null)
                {
                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.StyleIndex = percentIntegralFormatStyleIndex;
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                    if (percentIntegralCell.Value.HasValue)
                    {
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(DocumentFormat.OpenXml.DecimalValue.FromDecimal(percentIntegralCell.Value.Value));
                    }

                    row.AppendChild(cell);
                    continue;
                }

                percentFractionalCell = excelCell as ExcelCellPercentFractional;
                if (percentFractionalCell != null)
                {
                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.StyleIndex = percentFractionalFormatStyleIndex;
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                    if (percentFractionalCell.Value.HasValue)
                    {
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(DocumentFormat.OpenXml.DecimalValue.FromDecimal(percentFractionalCell.Value.Value));
                    }

                    row.AppendChild(cell);
                    continue;
                }

                dateCell = excelCell as ExcelCellDate;
                if (dateCell != null)
                {
                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.StyleIndex = dateFormatStyleIndex;
                    cell.DataType = null;
                    if (dateCell.Value.HasValue)
                    {
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(DocumentFormat.OpenXml.DoubleValue.FromDouble(dateCell.Value.Value.ToOADate()));
                    }

                    row.AppendChild(cell);
                    continue;
                }

                dateAndTimeCell = excelCell as ExcelCellDateAndTime;
                if (dateAndTimeCell != null)
                {
                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.StyleIndex = dateTimeFormatStyleIndex;
                    cell.DataType = null;
                    if (dateAndTimeCell.Value.HasValue)
                    {
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(DocumentFormat.OpenXml.DoubleValue.FromDouble(dateAndTimeCell.Value.Value.ToOADate()));
                    }

                    row.AppendChild(cell);
                    continue;
                }
            }

            sheetData.AppendChild(row);
        }

        /// <summary>
        /// Frees associated resources.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (spreadsheetDocument != null)
                {
                    spreadsheetDocument.Close();
                    spreadsheetDocument.Dispose();
                    spreadsheetDocument = null;
                }
            }
        }        
    }
}
