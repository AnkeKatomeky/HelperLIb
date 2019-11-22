using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HelperLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace HelperLibrary.ExcelOpenXml
{
    /// <summary>
    /// Writes data to Excel workbooks via OpenXML (uses SAX, supports large files).
    /// </summary>
    public class ExcelOpenXmlSaxWriter : IDisposable
    {
        private SpreadsheetDocument spreadsheetDocument;
        private OpenXmlWriter writer;

        private int currentRowIndex, columnsCount;
        private ExcelOpenXmlSaxWriterOptions options;

        /// <summary>
        /// Performs schema validation of Excel file.
        /// </summary>
        /// <param name="fileName">Excel file name.</param>
        /// <returns>True if file is valid, False otherwise.</returns>
        public static bool ValidateFile(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                DocumentFormat.OpenXml.Validation.OpenXmlValidator validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();

                int count = 0;
                foreach (DocumentFormat.OpenXml.Validation.ValidationErrorInfo error in validator.Validate(spreadsheetDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("ErrorType: " + error.ErrorType);
                    Console.WriteLine("Node: " + error.Node);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }

                return count == 0;
            }
        }

        /// <summary>
        /// Frees associated resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }

        /// <summary>
        /// Creates empty workbook (for reporting).
        /// </summary>
        /// <param name="fileName">Workbook file name.</param>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="columns">Array of column definitions.</param>
        /// <returns>True if success.</returns>
        public bool CreateWorkbookForReport(string fileName, string sheetName, ExcelColumn[] columns, ExcelOpenXmlSaxWriterOptions options)
        {
            try
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(fileName));
            }
            catch
            {
                return false;
            }
            spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook, true);

            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();

            workbookPart.Workbook = new Workbook();

            workbookPart.Workbook.BookViews = new BookViews();

            WorkbookView workbookView = new WorkbookView();

            workbookPart.Workbook.BookViews.Append(workbookView);

            workbookPart.Workbook.Sheets = new Sheets();

            this.options = options;

            AddDefaultStyles();

            AddStylesForReport();

            AddWorksheetForReport(sheetName, columns, options);

            return true;
        }

        /// <summary>
        /// Creates empty workbook (for reporting).
        /// </summary>
        /// <param name="fileName">Workbook file name.</param>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="columns">Array of column definitions.</param>
        /// <returns>True if success.</returns>
        public bool OpenWorkbookForReport(string fileName, ExcelOpenXmlSaxWriterOptions options)
        {

            if (!File.Exists(fileName))
            {
                return false;
            }

            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                spreadsheetDocument = SpreadsheetDocument.Open(fs, true);
            }

            //this.spreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(fileName, true);

            this.options = options;

            AddDefaultStyles();

            AddStylesForReport();

            return true;
        }

        /// <summary>
        /// Adds a new worksheet to current workbook (for reporting).
        /// </summary>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="columns">Array of column definitions.</param>
        /// <returns>True if success.</returns>
        public bool AddWorksheetForReport(string sheetName, ExcelColumn[] columns, ExcelOpenXmlSaxWriterOptions options)
        {
            if (spreadsheetDocument == null)
            {
                return false;
            }

            FinalizeWorksheet();

            currentRowIndex = 1;
            this.options = options;

            uint newSheetId = (uint)(spreadsheetDocument.WorkbookPart.Workbook.Sheets.Count() + 1);

            WorksheetPart worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = newSheetId,
                Name = sheetName
            };

            spreadsheetDocument.WorkbookPart.Workbook.Sheets.Append(sheet);
            Worksheet worksheet = new Worksheet();
            writer = OpenXmlWriter.Create(worksheetPart);
            writer.WriteStartElement(worksheet);

            #region Sheet dimension

            Dimension sheetDimension = new Dimension();

            sheetDimension.SetAttribute(new OpenXmlAttribute()
            {
                LocalName = "ref",
                Value = "A1:" + ExcelOpenXmlSaxReader.GetCellReference(1, columns.Length)
            });

            writer.WriteElement(sheetDimension);

            #endregion

            #region Sheet views

            SheetViews sheetViews = new SheetViews();

            #region Fix top row

            SheetView sheetView = new SheetView();
            if (newSheetId == 1)
            {
                sheetView.TabSelected = true;
            }

            sheetView.WorkbookViewId = 0;

            if (this.options.FixTopRow)
            {
                DocumentFormat.OpenXml.Spreadsheet.Selection selection1 = new DocumentFormat.OpenXml.Spreadsheet.Selection()
                {
                    Pane = PaneValues.BottomLeft
                };

                Pane pane1 = new Pane()
                {
                    VerticalSplit = 1d,
                    TopLeftCell = "A2",
                    ActivePane = PaneValues.BottomLeft,
                    State = PaneStateValues.Frozen
                };

                sheetView.Append(pane1);
                sheetView.Append(selection1);
            }
            if (this.options.FixLeftColums)
            {
                DocumentFormat.OpenXml.Spreadsheet.Selection selection1 = new DocumentFormat.OpenXml.Spreadsheet.Selection()
                {
                    Pane = PaneValues.BottomLeft
                };

                Pane pane1 = new Pane()
                {
                    HorizontalSplit = this.options.Offset,
                    VerticalSplit = 1d,
                    TopLeftCell = this.options.LockColumn,
                    ActivePane = PaneValues.BottomLeft,
                    State = PaneStateValues.Frozen
                };

                sheetView.Append(pane1);
                sheetView.Append(selection1);
            }

            sheetViews.Append(sheetView);

            #endregion

            writer.WriteElement(sheetViews);

            #endregion

            #region Columns

            writer.WriteStartElement(new Columns());

            Column column;
            ExcelColumn excelColumn;
            for (int i = 0; i < columns.Length; i++)
            {
                excelColumn = columns[i];

                column = new Column();

                column.Min = (uint)(i + 1);
                column.Max = (uint)(i + 1);
                column.CustomWidth = true;
                column.Width = excelColumn.Width;

                if (!string.IsNullOrEmpty(excelColumn.Group))
                {
                    column.SetAttribute(new OpenXmlAttribute("outlineLevel", string.Empty, excelColumn.Group));
                }

                writer.WriteElement(column);
            }

            writer.WriteEndElement();

            #endregion

            writer.WriteStartElement(new SheetData());

            AddHeaderToReport(columns, options.HeaderHeight);
            //if (options.IsToAddValidator)
            //{
            //    CreateValidator(worksheetPart, options.ValidatorParams, options.ValidatorColumn);
            //}

            return true;
        }

        /// <summary>
        /// Adds a new row with specified cells to current worksheet.
        /// </summary>
        /// <param name="cells">Array of cell definitions.</param>
        public void AddRowToReport(ExcelCell[] cells)
        {
            if (writer == null)
            {
                return;
            }

            if (cells.Length > columnsCount)
            {
                throw new ArgumentOutOfRangeException("Count of row's cells exceeds count of header's columns");
            }
            Row row = new Row();

            row.RowIndex = (uint)currentRowIndex;

            writer.WriteStartElement(row);

            Cell cell;

            ExcelCellString stringCell;
            ExcelCellBoolean booleanCell;
            ExcelCellNumberIntegral numberIntegralCell;
            ExcelCellNumberFractional numberFractionalCell;
            ExcelCellNumberFractionalLong numberFractionalLongCell;
            ExcelCellPercentIntegral percentIntegralCell;
            ExcelCellPercentFractional percentFractionalCell;
            ExcelCellDate dateCell;
            ExcelCellDateAndTime dateAndTimeCell;

            int currentCellIndex = 1;


            foreach (ExcelCell excelCell in cells)
            {
                #region string
                stringCell = excelCell as ExcelCellString;
                if (stringCell != null)
                {
                    cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(stringCell.Value);
                    cell.StyleIndex = options.DefaultString.StyleIndex;
                    if (!string.IsNullOrEmpty(stringCell.Formula))
                    {
                        cell.CellFormula = new CellFormula(stringCell.Formula);
                    }
                    if (stringCell.WordWrap)
                    {
                        cell.StyleIndex = options.WrapString.StyleIndex;
                    }
                    else
                    {
                        cell.StyleIndex = options.DefaultString.StyleIndex;
                    }
                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.DefaultB:
                                        cell.StyleIndex = options.DefaultStringBorder.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.PurpleString.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.GreenString.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.YellowString.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.RedString.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.BlueString.StyleIndex;
                                        break;
                                    case CellColor.BlueB:
                                        cell.StyleIndex = options.BlueStringBorder.StyleIndex;
                                        break;
                                    case CellColor.Marsh:
                                        cell.StyleIndex = options.MarshString.StyleIndex;
                                        break;
                                    case CellColor.Brown:
                                        cell.StyleIndex = options.BrownString.StyleIndex;
                                        break;
                                    case CellColor.Pastel:
                                        cell.StyleIndex = options.PastelString.StyleIndex;
                                        break;
                                    case CellColor.LGreen:
                                        cell.StyleIndex = options.LGreenString.StyleIndex;
                                        break;
                                    case CellColor.Pink:
                                        cell.StyleIndex = options.PinkString.StyleIndex;
                                        break;
                                    case CellColor.LBrown:
                                        cell.StyleIndex = options.LBrownString.StyleIndex;
                                        break;
                                    case CellColor.GrayB:
                                        cell.StyleIndex = options.GrayStringBorder.StyleIndex;
                                        break;
                                    case CellColor.PaleGreen:
                                        cell.StyleIndex = options.PaleGreenString.StyleIndex;
                                        break;
                                    case CellColor.PaleRed:
                                        cell.StyleIndex = options.PaleRedString.StyleIndex;
                                        break;
                                    case CellColor.WBlue:
                                        cell.StyleIndex = options.WBlueString.StyleIndex;
                                        break;
                                    case CellColor.DGray:
                                        cell.StyleIndex = options.DGrayString.StyleIndex;
                                        break;
                                    case CellColor.DBlue:
                                        cell.StyleIndex = options.DBlueString.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }
                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);

                    continue;
                }
                #endregion

                #region bool
                booleanCell = excelCell as ExcelCellBoolean;
                if (booleanCell != null)
                {
                    cell = new Cell();
                    cell.StyleIndex = options.Boolean.StyleIndex;
                    cell.DataType = CellValues.String;
                    if (booleanCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(booleanCell.Value.Value ? "Да" : null);
                    }
                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);
                    continue;
                }
                #endregion

                #region int
                numberIntegralCell = excelCell as ExcelCellNumberIntegral;
                if (numberIntegralCell != null)
                {
                    cell = new Cell();
                    cell.StyleIndex = options.DefaultInt.StyleIndex;
                    cell.DataType = CellValues.Number;
                    if (!string.IsNullOrEmpty(numberIntegralCell.Formula))
                    {
                        cell.CellFormula = new CellFormula(numberIntegralCell.Formula);
                    }
                    if (numberIntegralCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(Int32Value.FromInt32(numberIntegralCell.Value.Value));
                    }
                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.DefaultInt.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.PurpleInt.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.GreenInt.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.YellowInt.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.RedInt.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.BlueInt.StyleIndex;
                                        break;
                                    case CellColor.BlueB:
                                        cell.StyleIndex = options.BlueIntBorder.StyleIndex;
                                        break;
                                    case CellColor.DefaultB:
                                        cell.StyleIndex = options.DefaultIntBorder.StyleIndex;
                                        break;
                                    case CellColor.PaleGreen:
                                        cell.StyleIndex = options.PaleGreenInt.StyleIndex;
                                        break;
                                    case CellColor.PaleRed:
                                        cell.StyleIndex = options.PaleRedInt.StyleIndex;
                                        break;
                                    case CellColor.WBlue:
                                        cell.StyleIndex = options.WBlueInt.StyleIndex;
                                        break;
                                    case CellColor.GrayB:
                                        cell.StyleIndex = options.GrayIntBorder.StyleIndex;
                                        break;
                                    case CellColor.Gray:
                                        cell.StyleIndex = options.GrayInt.StyleIndex;
                                        break;
                                    case CellColor.DGray:
                                        cell.StyleIndex = options.DGrayInt.StyleIndex;
                                        break;
                                    case CellColor.DBlue:
                                        cell.StyleIndex = options.DBlueInt.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);
                    continue;
                }
                #endregion

                #region fract
                numberFractionalCell = excelCell as ExcelCellNumberFractional;
                if (numberFractionalCell != null)
                {
                    cell = new Cell();
                    cell.StyleIndex = options.DefaultFract.StyleIndex;
                    cell.DataType = CellValues.Number;
                    if (!string.IsNullOrEmpty(numberFractionalCell.Formula))
                    {
                        cell.CellFormula = new CellFormula(numberFractionalCell.Formula);
                    }
                    if (numberFractionalCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(DecimalValue.FromDecimal(numberFractionalCell.Value.Value));
                    }
                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.DefaultFract.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.PurpleFract.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.GreenFract.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.YellowFract.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.RedFract.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.BlueFract.StyleIndex;
                                        break;
                                    case CellColor.BlueB:
                                        cell.StyleIndex = options.BlueFractBorder.StyleIndex;
                                        break;
                                    case CellColor.DefaultB:
                                        cell.StyleIndex = options.DefaultFractBorder.StyleIndex;
                                        break;
                                    case CellColor.GreenB:
                                        cell.StyleIndex = options.GreenFractBorder.StyleIndex;
                                        break;
                                    case CellColor.YellowB:
                                        cell.StyleIndex = options.YellowFractBorder.StyleIndex;
                                        break;
                                    case CellColor.RedB:
                                        cell.StyleIndex = options.RedFractBorder.StyleIndex;
                                        break;
                                    case CellColor.PaleGreen:
                                        cell.StyleIndex = options.PaleGreenFract.StyleIndex;
                                        break;
                                    case CellColor.PaleRed:
                                        cell.StyleIndex = options.PaleRedFract.StyleIndex;
                                        break;
                                    case CellColor.GrayB:
                                        cell.StyleIndex = options.GrayFractBorder.StyleIndex;
                                        break;
                                    case CellColor.Gray:
                                        cell.StyleIndex = options.GrayFract.StyleIndex;
                                        break;
                                    case CellColor.DGray:
                                        cell.StyleIndex = options.DGrayFract.StyleIndex;
                                        break;
                                    case CellColor.DBlue:
                                        cell.StyleIndex = options.DBlueFract.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.DefaultFract.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);
                    continue;
                }
                #endregion

                #region fractLong
                numberFractionalLongCell = excelCell as ExcelCellNumberFractionalLong;
                if (numberFractionalLongCell != null)
                {
                    cell = new Cell();
                    cell.StyleIndex = options.DefaultFractLong.StyleIndex;
                    cell.DataType = CellValues.Number;
                    if (!string.IsNullOrEmpty(numberFractionalLongCell.Formula))
                    {
                        cell.CellFormula = new CellFormula(numberFractionalLongCell.Formula);
                    }
                    if (numberFractionalLongCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(DecimalValue.FromDecimal(numberFractionalLongCell.Value.Value));
                    }
                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.DefaultFract.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.PurpleFract.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.GreenFract.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.YellowFract.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.RedFract.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.BlueFract.StyleIndex;
                                        break;
                                    case CellColor.BlueB:
                                        cell.StyleIndex = options.BlueFractBorder.StyleIndex;
                                        break;
                                    case CellColor.DefaultB:
                                        cell.StyleIndex = options.DefaultFractBorder.StyleIndex;
                                        break;
                                    case CellColor.GreenB:
                                        cell.StyleIndex = options.GreenFractBorder.StyleIndex;
                                        break;
                                    case CellColor.YellowB:
                                        cell.StyleIndex = options.YellowFractBorder.StyleIndex;
                                        break;
                                    case CellColor.RedB:
                                        cell.StyleIndex = options.RedFractBorder.StyleIndex;
                                        break;
                                    case CellColor.PaleGreen:
                                        cell.StyleIndex = options.PaleGreenFract.StyleIndex;
                                        break;
                                    case CellColor.PaleRed:
                                        cell.StyleIndex = options.PaleRedFract.StyleIndex;
                                        break;
                                    case CellColor.GrayB:
                                        cell.StyleIndex = options.GrayFractBorder.StyleIndex;
                                        break;
                                    case CellColor.Gray:
                                        cell.StyleIndex = options.GrayFract.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.DefaultFract.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);
                    continue;
                }
                #endregion

                #region fractLong
                numberFractionalLongCell = excelCell as ExcelCellNumberFractionalLong;
                if (numberFractionalLongCell != null)
                {
                    cell = new Cell
                    {
                        StyleIndex = options.DefaultFractLong.StyleIndex,
                        DataType = CellValues.Number
                    };
                    if (!string.IsNullOrEmpty(numberFractionalLongCell.Formula))
                    {
                        cell.CellFormula = new CellFormula(numberFractionalLongCell.Formula);
                    }
                    if (numberFractionalLongCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(DecimalValue.FromDecimal(numberFractionalLongCell.Value.Value));
                    }
                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.DefaultFract.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.PurpleFract.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.GreenFract.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.YellowFract.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.RedFract.StyleIndex;
                                        break;
                                    case CellColor.PaleRed:
                                        cell.StyleIndex = options.PaleRedFract.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.BlueFract.StyleIndex;
                                        break;
                                    case CellColor.BlueB:
                                        cell.StyleIndex = options.BlueFractBorder.StyleIndex;
                                        break;

                                    case CellColor.DefaultB:
                                        cell.StyleIndex = options.DefaultFractBorder.StyleIndex;
                                        break;
                                    case CellColor.GreenB:
                                        cell.StyleIndex = options.GreenFractBorder.StyleIndex;
                                        break;
                                    case CellColor.YellowB:
                                        cell.StyleIndex = options.YellowFractBorder.StyleIndex;
                                        break;
                                    case CellColor.RedB:
                                        cell.StyleIndex = options.RedFractBorder.StyleIndex;
                                        break;

                                    default:
                                        cell.StyleIndex = options.DefaultFract.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }

                    writer.WriteElement(cell);
                    continue;
                }
                #endregion

                #region percents int
                percentIntegralCell = excelCell as ExcelCellPercentIntegral;
                if (percentIntegralCell != null)
                {
                    cell = new Cell();
                    cell.StyleIndex = options.PercentDefaultInt.StyleIndex;
                    cell.DataType = CellValues.Number;
                    if (percentIntegralCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(DecimalValue.FromDecimal(percentIntegralCell.Value.Value));
                    }

                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.PercentDefaultInt.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.PercentPurpleInt.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.PercentGreenInt.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.PercentYellowInt.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.PercentRedInt.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.PercentBlueInt.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.PercentDefaultInt.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);
                    continue;
                }
                #endregion

                #region percents fract
                percentFractionalCell = excelCell as ExcelCellPercentFractional;
                if (percentFractionalCell != null)
                {
                    cell = new Cell();
                    cell.StyleIndex = options.PercentDefaultFract.StyleIndex;
                    cell.DataType = CellValues.Number;
                    if (percentFractionalCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(DecimalValue.FromDecimal(percentFractionalCell.Value.Value));
                    }
                    if (excelCell.UseErrorStyle != null)
                    {

                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.PercentDefaultFract.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.PercentPurpleFract.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.PercentGreenFract.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.PercentYellowFract.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.PercentRedFract.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.PercentBlueFract.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.PercentDefaultFract.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);
                    continue;
                }
                #endregion

                #region date
                dateCell = excelCell as ExcelCellDate;
                if (dateCell != null)
                {
                    cell = new Cell();
                    cell.StyleIndex = options.Date.StyleIndex;
                    cell.DataType = null;
                    if (dateCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(DoubleValue.FromDouble(dateCell.Value.Value.ToOADate()));
                    }
                    if (!string.IsNullOrEmpty(dateCell.Formula))
                    {
                        cell.CellFormula = new CellFormula(dateCell.Formula);
                    }

                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.Date.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.PurpleDate.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.GreenDate.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.YellowDate.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.RedDate.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.BlueDate.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.Date.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);
                    continue;
                }
                #endregion

                #region date an time
                dateAndTimeCell = excelCell as ExcelCellDateAndTime;
                if (dateAndTimeCell != null)
                {
                    cell = new Cell();
                    cell.StyleIndex = options.DateAndTime.StyleIndex;
                    cell.DataType = null;
                    if (dateAndTimeCell.Value.HasValue)
                    {
                        cell.CellValue = new CellValue(DoubleValue.FromDouble(dateAndTimeCell.Value.Value.ToOADate()));
                    }
                    if (!string.IsNullOrEmpty(dateAndTimeCell.Formula))
                    {
                        cell.CellFormula = new CellFormula(dateAndTimeCell.Formula);
                    }

                    if (excelCell.UseErrorStyle != null)
                    {
                        foreach (IfFormat item in excelCell.UseErrorStyle)
                        {
                            if (item.IfCondition)
                            {
                                switch (item.ColorCondition)
                                {
                                    case CellColor.Default:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Purple:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Green:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Yellow:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Red:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    case CellColor.Blue:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                    default:
                                        cell.StyleIndex = options.DefaultString.StyleIndex;
                                        break;
                                }
                            }
                        }
                    }

                    cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                    if (excelCell.Reference != null)
                    {
                        excelCell.Reference.cellRef = cell.CellReference.HasValue ? cell.CellReference.Value : "";
                    }
                    writer.WriteElement(cell);
                    continue;
                }
                #endregion


            }
            writer.WriteEndElement();

            currentRowIndex++;
        }

        /// <summary>
        /// Closes current workbook.
        /// </summary>
        public void CloseWorkbookForReport()
        {
            FinalizeWorksheet();

            SaveWorkbook();

            CleanupResources();
        }

        /// <summary>
        /// Frees associated resources.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                CleanupResources();
            }
        }

        private void CleanupResources()
        {
            if (writer != null)
            {
                writer.Close();
                writer.Dispose();
                writer = null;
            }

            if (spreadsheetDocument != null)
            {
                spreadsheetDocument.Close();
                spreadsheetDocument.Dispose();
                spreadsheetDocument = null;
            }
        }

        private void SaveWorkbook()
        {
            if (spreadsheetDocument != null)
            {
                spreadsheetDocument.WorkbookPart.Workbook.Save();
            }
        }

        private void FinalizeWorksheet()
        {
            if (writer != null)
            {
                // End of SheetData
                writer.WriteEndElement();

                #region Auto filter

                if (options.AddAutoFilter)
                {
                    AutoFilter autoFilter = new AutoFilter()
                    {
                        Reference = ExcelOpenXmlSaxReader.GetCellReference(1, 1) + ":" + ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex - 1, columnsCount)
                    };

                    writer.WriteElement(autoFilter);
                }

                #endregion

                // End of worksheet
                writer.WriteEndElement();

                writer.Close();
                writer.Dispose();
                writer = null;
            }
        }

        private void AddStylesForReport()
        {
            if (spreadsheetDocument == null)
            {
                return;
            }
            if (spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats == null)
            {
                spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats = new CellFormats();
            }

            if (spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats == null)
            {
                spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
            }
            foreach (StyleExcel item in options.Styles)
            {
                item.SetStyle(spreadsheetDocument.WorkbookPart.WorkbookStylesPart, spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++);
            }

            #region old
            //DocumentFormat.OpenXml.Spreadsheet.CellFormat cellFormat;
            //DocumentFormat.OpenXml.Spreadsheet.NumberingFormat numberingFormat;

            //DocumentFormat.OpenXml.UInt32Value customFormatId = 164;

            //#region Header text

            ///////////////////DEFAULT 
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.FontId = 1;
            //cellFormat.ApplyFont = true;
            //cellFormat.FillId = 2;
            //cellFormat.ApplyFill = true;
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
            //    WrapText = true
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.defaultHeaderFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            ///////////////////PURPLE   
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.FontId = 2;
            //cellFormat.ApplyFont = true;
            //cellFormat.FillId = 3;
            //cellFormat.ApplyFill = true;
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
            //    WrapText = true
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.purpleHeaderFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            ////////////////////GREEN
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.FontId = 3;
            //cellFormat.ApplyFont = true;
            //cellFormat.FillId = 4;
            //cellFormat.ApplyFill = true;
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
            //    WrapText = true
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.greenHeaderFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //////////////////////YELOW
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.FontId = 4;
            //cellFormat.ApplyFont = true;
            //cellFormat.FillId = 5;
            //cellFormat.ApplyFill = true;
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
            //    WrapText = true
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.yellowHeaderFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            ////////////////////RED
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.FontId = 5;
            //cellFormat.ApplyFont = true;
            //cellFormat.FillId = 6;
            //cellFormat.ApplyFill = true;
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
            //    WrapText = true
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.redHeaderFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            ////////////////////BLUE  
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.FontId = 6;
            //cellFormat.ApplyFont = true;
            //cellFormat.FillId = 7;
            //cellFormat.ApplyFill = true;
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
            //    WrapText = true
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.blueHeaderFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            //#endregion

            //#region String

            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.stringFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion

            //#region String with word wrap

            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top,
            //    WrapText = true
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.stringFormatStyleIndex2 = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion

            //#region Number (integral)
            ////Default
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 3; // #,##0
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.numberIntegralFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            ////Purple
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 3; // #,##0
            //cellFormat.FillId = 3;
            //cellFormat.ApplyFill = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.purpleNumberIntegralFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            ////Green
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 3; // #,##0
            //cellFormat.FillId = 4;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.greenNumberIntegralFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            ////Yellow
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 3; // #,##0
            //cellFormat.FillId = 5;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.yellowNumberIntegralFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            ////RED
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 3; // #,##0
            //cellFormat.FillId = 6;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.redNumberIntegralFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            ////BLUE
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 3; // #,##0
            //cellFormat.FillId = 7;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.blueNumberIntegralFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion

            //#region Number (fractional)

            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 4; // #,##0.00
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.numberFractionalFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            ////Purple
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 4; // #,##0.00
            //cellFormat.FillId = 3;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.purpleNumberFractionalFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            ////Green
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 4; // #,##0.00
            //cellFormat.FillId = 4;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.greenNumberFractionalFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            ////Yellow
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 4; // #,##0.00
            //cellFormat.FillId = 5;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.yellowNumberFractionalFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            ////Red
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 4; // #,##0.00
            //cellFormat.FillId = 6;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.redNumberFractionalFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            ////Blue
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 4; // #,##0.00
            //cellFormat.FillId = 7;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.blueNumberFractionalFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion

            //#region Percent (integral)

            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 9; // 0%
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.percentIntegralFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion

            //#region Percent (fractional)
            //////DEFAULT
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = 10; // 0.00%
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.percentFractionalFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;
            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            ////YELLOW
            //cellFormat.NumberFormatId = 10; // 0.00%
            //cellFormat.FillId = 5;
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.yellowPercentFractionalFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion

            //#region Date

            //numberingFormat = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat();
            //numberingFormat.NumberFormatId = customFormatId++;
            //numberingFormat.FormatCode = DocumentFormat.OpenXml.StringValue.FromString("dd.mm.yyyy");

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>(numberingFormat);

            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = numberingFormat.NumberFormatId; // 14; // mm-dd-yy
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.dateFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion

            //#region Date and time

            //numberingFormat = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat();
            //numberingFormat.NumberFormatId = customFormatId++;
            //numberingFormat.FormatCode = DocumentFormat.OpenXml.StringValue.FromString("dd.mm.yyyy hh:mm");

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>(numberingFormat);

            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.NumberFormatId = numberingFormat.NumberFormatId; // 22; // m/d/yy h:mm
            //cellFormat.ApplyNumberFormat = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.dateTimeFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion

            //#region Boolean

            //cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat();
            //cellFormat.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Alignment
            //{
            //    Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            //    Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top
            //});

            //this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
            //this.booleanFormatStyleIndex = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count++;

            //#endregion
            #endregion
        }

        private void AddDefaultStyles()
        {
            if (spreadsheetDocument == null)
            {
                return;
            }

            WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = new Stylesheet();

            #region Fonts
            workbookStylesPart.Stylesheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();
            workbookStylesPart.Stylesheet.Fonts.AppendChild(new Font());
            workbookStylesPart.Stylesheet.Fonts.Count = 1;
            foreach (StyleFont item in options.Fonts)
            {
                item.SetFont(workbookStylesPart, workbookStylesPart.Stylesheet.Fonts.Count++);
            }

            #endregion

            #region Fills

            workbookStylesPart.Stylesheet.Fills = new Fills();

            workbookStylesPart.Stylesheet.Fills.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Fill
            {
                PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
                {
                    PatternType = PatternValues.None
                }
            }); // required, reserved by Excel

            workbookStylesPart.Stylesheet.Fills.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Fill
            {
                PatternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
                {
                    PatternType = PatternValues.Gray125
                }
            }); // required, reserved by Excel
            workbookStylesPart.Stylesheet.Fills.Count = 2;
            foreach (StyleFill item in options.Fills)
            {
                item.SetFill(workbookStylesPart, workbookStylesPart.Stylesheet.Fills.Count++);
            }
            #endregion

            #region Borders

            workbookStylesPart.Stylesheet.Borders = new Borders();
            workbookStylesPart.Stylesheet.Borders.AppendChild(new Border());

            Border border2 = new Border();

            DocumentFormat.OpenXml.Spreadsheet.LeftBorder leftBorder2 = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder() { Style = BorderStyleValues.Thin };
            Color color1 = new Color() { Indexed = 64U };

            leftBorder2.Append(color1);

            DocumentFormat.OpenXml.Spreadsheet.RightBorder rightBorder2 = new DocumentFormat.OpenXml.Spreadsheet.RightBorder() { Style = BorderStyleValues.Thin };
            Color color2 = new Color() { Indexed = 64U };

            rightBorder2.Append(color2);

            DocumentFormat.OpenXml.Spreadsheet.TopBorder topBorder2 = new DocumentFormat.OpenXml.Spreadsheet.TopBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = 64U };

            topBorder2.Append(color3);

            DocumentFormat.OpenXml.Spreadsheet.BottomBorder bottomBorder2 = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = 64U };

            bottomBorder2.Append(color4);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);
            workbookStylesPart.Stylesheet.Borders.AppendChild(border2);

            workbookStylesPart.Stylesheet.Borders.Count = 2;

            #endregion

            #region Cell style formats

            workbookStylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            workbookStylesPart.Stylesheet.CellStyleFormats.Count = 1;
            workbookStylesPart.Stylesheet.CellStyleFormats.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.CellFormat());

            #endregion

            #region Cell formats

            workbookStylesPart.Stylesheet.CellFormats = new CellFormats();

            // empty one for index 0, seems to be required
            workbookStylesPart.Stylesheet.CellFormats.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.CellFormat());

            // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
            workbookStylesPart.Stylesheet.CellFormats.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.CellFormat
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 2,
                ApplyFill = true
            }).AppendChild(new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center
            });

            workbookStylesPart.Stylesheet.CellFormats.Count = 2;

            #endregion

            workbookStylesPart.Stylesheet.Save();
        }

        private void AddHeaderToReport(ExcelColumn[] columns, int headerHeight)
        {
            if (writer == null)
            {
                return;
            }

            Row row = new Row();
            if (headerHeight > 0)
            {
                row.Height = headerHeight;
                row.CustomHeight = true;
            }

            row.RowIndex = (uint)currentRowIndex;

            writer.WriteStartElement(row, row.GetAttributes());

            Cell cell;

            int currentCellIndex = 1;

            foreach (ExcelColumn column in columns)
            {
                cell = new Cell();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(column.Title);

                switch (column.Style)
                {
                    case CellColor.Default:
                        cell.StyleIndex = options.DefaultHeader.StyleIndex;
                        break;
                    case CellColor.Purple:
                        cell.StyleIndex = options.PurpleHeader.StyleIndex;
                        break;
                    case CellColor.Green:
                        cell.StyleIndex = options.GreenHeader.StyleIndex;
                        break;
                    case CellColor.Yellow:
                        cell.StyleIndex = options.YellowHeader.StyleIndex;
                        break;
                    case CellColor.Red:
                        cell.StyleIndex = options.RedHeader.StyleIndex;
                        break;
                    case CellColor.Blue:
                        cell.StyleIndex = options.BlueHeader.StyleIndex;
                        break;
                    case CellColor.DBlue:
                        cell.StyleIndex = options.DBlueHeader.StyleIndex;
                        break;
                    case CellColor.Marsh:
                        cell.StyleIndex = options.MarshHeader.StyleIndex;
                        break;
                    case CellColor.Brown:
                        cell.StyleIndex = options.BrownHeader.StyleIndex;
                        break;
                    case CellColor.Pastel:
                        cell.StyleIndex = options.PastelHeader.StyleIndex;
                        break;
                    case CellColor.LGreen:
                        cell.StyleIndex = options.LGreenHeader.StyleIndex;
                        break;
                    case CellColor.Pink:
                        cell.StyleIndex = options.PinkHeader.StyleIndex;
                        break;
                    case CellColor.LBrown:
                        cell.StyleIndex = options.LBrownHeader.StyleIndex;
                        break;
                    case CellColor.DefaultB:
                        cell.StyleIndex = options.DefaultHeaderBorder.StyleIndex;
                        break;
                    case CellColor.GreenB:
                        cell.StyleIndex = options.GreenHeaderBorder.StyleIndex;
                        break;
                    case CellColor.YellowB:
                        cell.StyleIndex = options.YellowHeaderBorder.StyleIndex;
                        break;
                    case CellColor.RedB:
                        cell.StyleIndex = options.RedHeaderBorder.StyleIndex;
                        break;
                    case CellColor.WGreen:
                        cell.StyleIndex = options.WGreenHeader.StyleIndex;
                        break;
                    case CellColor.Cyian:
                        cell.StyleIndex = options.CyianHeader.StyleIndex;
                        break;
                    case CellColor.PaleGreen:
                        cell.StyleIndex = options.PaleGreenHeader.StyleIndex;
                        break;
                    case CellColor.PaleGreenAndRed:
                        cell.StyleIndex = options.PaleGreenAndRedHeader.StyleIndex;
                        break;
                    case CellColor.PaleRed:
                        cell.StyleIndex = options.PaleRedHeader.StyleIndex;
                        break;
                    case CellColor.BlueAndRed:
                        cell.StyleIndex = options.BlueAndRedHeader.StyleIndex;
                        break;
                    case CellColor.BlueB:
                        cell.StyleIndex = options.BlueHeaderBorder.StyleIndex;
                        break;
                    default:
                        cell.StyleIndex = options.DefaultHeader.StyleIndex;
                        break;


                }

                cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);

                writer.WriteElement(cell);
            }

            writer.WriteEndElement();

            currentRowIndex++;

            columnsCount = columns.Length;
        }

        public static void InsertChartInSpreadsheet(string docName, string worksheetName, ChartData data)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
                if (sheets.Count() == 0)
                {
                    // The specified worksheet does not exist.
                    return;
                }
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

                DrawingsPart drawingsPart;
                // Add a new drawing to the worksheet.
                if (worksheetPart.DrawingsPart == null)
                {
                    drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                    worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                    worksheetPart.Worksheet.Save();
                }
                else
                {
                    drawingsPart = worksheetPart.DrawingsPart;
                }


                // Add a new chart and set the chart language to English-US.
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new ChartSpace();
                chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
                DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Chart());

                // Create a new clustered column chart.
                PlotArea plotArea = chart.AppendChild(new PlotArea());
                Layout layout = plotArea.AppendChild(new Layout());
                BarChart barChart = plotArea.AppendChild(new BarChart(
                    new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                    new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

                uint i = 0;

                // Iterate through each key in the Dictionary collection and add the key to the chart Series
                // and add the corresponding value to the chart Values.
                foreach (string key in data.Data.Keys)
                {
                    BarChartSeries barChartSeries = barChart.AppendChild(new BarChartSeries(
                        new Index() { Val = new UInt32Value(i) },
                        new Order() { Val = new UInt32Value(i) },
                        new SeriesText(new NumericValue() { Text = key })));

                    // Adding category axis to the chart
                    CategoryAxisData categoryAxisData = barChartSeries.AppendChild(new CategoryAxisData());

                    string formulaCat = "";
                    if (data.CoresNumbers != null)
                    {
                        formulaCat = "(";
                        foreach (int item in data.CoresNumbers)
                        {
                            formulaCat += string.Format("Summary!${0}${1},", data.CoreColumn, item);
                        }
                        formulaCat = formulaCat.Remove(formulaCat.Length - 1);
                        formulaCat += ")";
                    }
                    else
                    {
                        formulaCat = string.Format("Summary!${2}${0}:${2}${1}", data.From, data.To, data.CoreColumn);
                    }
                    StringReference stringReference = categoryAxisData.AppendChild(new StringReference()
                    {
                        Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaCat }
                    });

                    string formulaVal = "";
                    if (data.CoresNumbers != null)
                    {
                        formulaVal = "(";
                        foreach (int item in data.CoresNumbers)
                        {
                            formulaVal += string.Format("Summary!${0}${1},", data.Data[key], item);
                        }
                        formulaVal = formulaVal.Remove(formulaVal.Length - 1);
                        formulaVal += ")";

                    }
                    else
                    {
                        formulaVal = string.Format("Summary!${0}${1}:${0}${2}", data.Data[key], data.From, data.To);
                    }

                    DocumentFormat.OpenXml.Drawing.Charts.Values values = barChartSeries.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());

                    NumberReference numberReference = values.AppendChild(new NumberReference()
                    {
                        Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaVal }
                    });

                    i++;
                }

                barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
                barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

                // Add the Category Axis.

                CategoryAxis catAx = plotArea.AppendChild(new CategoryAxis(
                    new AxisId() { Val = new UInt32Value(48650112u) },
                    new Scaling(new Orientation()
                    {
                        Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                    }),
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                    new CrossingAxis() { Val = new UInt32Value(48672768U) },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new AutoLabeled() { Val = new BooleanValue(true) },
                    new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                    new LabelOffset() { Val = new UInt16Value((ushort)100) }));

                // Add the Value Axis.
                ValueAxis valAx = plotArea.AppendChild(new ValueAxis(
                    new AxisId() { Val = new UInt32Value(48672768u) },
                    new Scaling(new Orientation()
                    {
                        Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                    }),
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                    new MajorGridlines(),
                    new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                    {
                        FormatCode = new StringValue("General"),
                        SourceLinked = new BooleanValue(true)
                    },
                    new TickLabelPosition()
                    {
                        Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo)
                    },
                    new CrossingAxis() { Val = new UInt32Value(48650112U) },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

                // Add the chart Legend.
                Legend legend = chart.AppendChild(new Legend(
                    new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
                    new Layout()));

                chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

                // Save the chart part.
                chartPart.ChartSpace.Save();

                // Position the chart on the worksheet using a TwoCellAnchor object.
                if (drawingsPart.WorksheetDrawing == null)
                {
                    drawingsPart.WorksheetDrawing = new WorksheetDrawing();
                }
                TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());
                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                    new ColumnId(data.Columm.ToString()),//new ColumnId("0"),
                    new ColumnOffset("0"),
                    new RowId(data.Row.ToString()),//new RowId("26"),
                    new RowOffset("0")));
                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                    new ColumnId((data.Width + data.Columm).ToString()),//new ColumnId("5"),
                    new ColumnOffset("0"),
                    new RowId((data.Row + data.Height).ToString()),//new RowId("40"),
                    new RowOffset("0")));

                // Append a GraphicFrame to the TwoCellAnchor object.
                DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                    twoCellAnchor.AppendChild(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());

                graphicFrame.Macro = "";

                graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = data.Title },
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

                graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L }, new Extents() { Cx = 0L, Cy = 0L }));

                graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

                twoCellAnchor.Append(new ClientData());

                // Save the WorksheetDrawing object.
                drawingsPart.WorksheetDrawing.Save();
            }

        }

        public static void CreateValidator(string docName, string worksheetName, string dataContainingSheet, string column)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
                if (sheets.Count() == 0)
                {
                    // The specified worksheet does not exist.
                    return;
                }
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);


                /***  DATA VALIDATION CODE ***/
                DataValidations dataValidations = new DataValidations();
                DataValidation dataValidation = new DataValidation
                {
                    Type = DataValidationValues.List,
                    AllowBlank = true,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Format("{0}1:{0}1048576", column) }
                };

                dataValidation.Append(
                    new Formula1 { Text = dataContainingSheet }
                    //new Formula1 { Text = "\"FirstChoice,SecondChoice,ThirdChoice\"" }
                    //new Formula1(string.Format("'{0}'!${1}:${1}", dataContainingSheet, column))
                    );
                dataValidations.Append(dataValidation);

                var wsp = worksheetPart;
                wsp.Worksheet.AppendChild(dataValidations);
            }
        }

        public static void InsertDataInSpreadsheet(string docName, string worksheetName, ExcelCell[] data, ExcelColumn[] header)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                UInt32Value headStyle;
                UInt32Value textStyle;
                UInt32Value digitStyle;
                UInt32Value dateStyle;

                #region Creating
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
                if (sheets.Count() == 0)
                {
                    // Add a blank WorksheetPart.
                    WorksheetPart newWorksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
                    newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                    Sheets sheets1 = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    string relationshipId = document.WorkbookPart.GetIdOfPart(newWorksheetPart);

                    // Get a unique ID for the new worksheet.
                    uint sheetId = 1;
                    if (sheets1.Elements<Sheet>().Count() > 0)
                    {
                        sheetId = sheets1.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    // Give the new worksheet a name.
                    string sheetName = worksheetName;

                    // Append the new worksheet and associate it with the workbook.
                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                    sheets1.Append(sheet);
                    sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
                }
                #endregion

                #region STYLES
                IEnumerable<Sheet> ordersheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Order");
                WorksheetPart orderworksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(ordersheets.First().Id);
                SheetData ordersheetData = orderworksheetPart.Worksheet.GetFirstChild<SheetData>();
                Row headerRow = ordersheetData.Elements<Row>().FirstOrDefault();
                Row mainRow = ordersheetData.Elements<Row>().LastOrDefault();

                Cell heads = (Cell)headerRow.FirstOrDefault();
                Cell text = (Cell)mainRow.Elements().Where(q => q.OuterXml.Contains("A")).FirstOrDefault();
                Cell digits = (Cell)mainRow.Elements().Where(q => q.OuterXml.Contains("I")).FirstOrDefault();
                Cell date = (Cell)mainRow.Elements().Where(q => q.OuterXml.Contains("M")).FirstOrDefault();

                headStyle = heads.StyleIndex;
                textStyle = text.StyleIndex;
                digitStyle = digits.StyleIndex;
                dateStyle = date.StyleIndex;

                //headStyle = uint.Parse(Regex.Match(Regex.Match(heads, "(s=\"[\\d]*\")").Value, "(\\d\\d)").Value);
                #endregion

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                Row lastRow = sheetData.Elements<Row>().LastOrDefault();
                Row row;
                Row headRow;
                int currentCellIndex = 1;
                int currentRowIndex;

                Cell cell;

                if (lastRow == null)
                {
                    headRow = new Row() { RowIndex = 1 };
                    sheetData.InsertAt(headRow, 0);
                    currentRowIndex = (int)headRow.RowIndex.Value;

                    //Columns columns = new Columns();
                    //Column col;
                    //for (int i = 0; i < header.Length; i++)
                    //{
                    //    col = new Column();

                    //    col.Min = (uint)(i + 1);
                    //    col.Max = (uint)(i + 1);
                    //    col.CustomWidth = true;
                    //    col.Width = header[i].Width;
                    //}
                    //worksheetPart.Worksheet.Append(columns);

                    foreach (ExcelColumn column in header)
                    {
                        cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(column.Title);
                        cell.StyleIndex = headStyle;
                        cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                        headRow.Append(cell);
                    }

                }
                row = new Row() { RowIndex = (lastRow.RowIndex + 1) };
                sheetData.InsertAfter(row, lastRow);


                if (true)
                {
                    ExcelCellString stringCell;
                    ExcelCellBoolean booleanCell;
                    ExcelCellNumberIntegral numberIntegralCell;
                    ExcelCellNumberFractional numberFractionalCell;
                    ExcelCellPercentIntegral percentIntegralCell;
                    ExcelCellPercentFractional percentFractionalCell;
                    ExcelCellDate dateCell;
                    ExcelCellDateAndTime dateAndTimeCell;

                    currentCellIndex = 1;
                    currentRowIndex = (int)row.RowIndex.Value;
                    foreach (ExcelCell excelCell in data)
                    {
                        #region string
                        stringCell = excelCell as ExcelCellString;
                        if (stringCell != null)
                        {
                            cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(stringCell.Value);
                            cell.StyleIndex = textStyle;

                            cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                            row.Append(cell);

                            continue;
                        }
                        #endregion

                        #region bool
                        booleanCell = excelCell as ExcelCellBoolean;
                        if (booleanCell != null)
                        {
                            cell = new Cell();
                            cell.StyleIndex = new UInt32Value((uint)34);
                            cell.DataType = CellValues.String;
                            if (booleanCell.Value.HasValue)
                            {
                                cell.CellValue = new CellValue(booleanCell.Value.Value ? "Да" : null);
                            }

                            cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                            row.Append(cell);

                            continue;
                        }
                        #endregion

                        #region int
                        numberIntegralCell = excelCell as ExcelCellNumberIntegral;
                        if (numberIntegralCell != null)
                        {
                            cell = new Cell();
                            cell.StyleIndex = digitStyle;
                            cell.DataType = CellValues.Number;
                            if (!string.IsNullOrEmpty(numberIntegralCell.Formula))
                            {
                                cell.CellFormula = new CellFormula(numberIntegralCell.Formula);
                            }
                            if (numberIntegralCell.Value.HasValue)
                            {
                                cell.CellValue = new CellValue(Int32Value.FromInt32(numberIntegralCell.Value.Value));
                            }

                            cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                            row.Append(cell);

                            continue;
                        }
                        #endregion

                        #region fract
                        numberFractionalCell = excelCell as ExcelCellNumberFractional;
                        if (numberFractionalCell != null)
                        {
                            cell = new Cell();
                            cell.StyleIndex = digitStyle;
                            cell.DataType = CellValues.Number;
                            if (!string.IsNullOrEmpty(numberFractionalCell.Formula))
                            {
                                cell.CellFormula = new CellFormula(numberFractionalCell.Formula);
                            }
                            if (numberFractionalCell.Value.HasValue)
                            {
                                cell.CellValue = new CellValue(DecimalValue.FromDecimal(numberFractionalCell.Value.Value));
                            }

                            cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                            row.Append(cell);

                            continue;
                        }
                        #endregion

                        #region percents int
                        percentIntegralCell = excelCell as ExcelCellPercentIntegral;
                        if (percentIntegralCell != null)
                        {
                            cell = new Cell();
                            cell.StyleIndex = digitStyle;
                            cell.DataType = CellValues.Number;
                            if (percentIntegralCell.Value.HasValue)
                            {
                                cell.CellValue = new CellValue(DecimalValue.FromDecimal(percentIntegralCell.Value.Value));
                            }
                            cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                            row.Append(cell);

                            continue;
                        }
                        #endregion

                        #region percents fract
                        percentFractionalCell = excelCell as ExcelCellPercentFractional;
                        if (percentFractionalCell != null)
                        {
                            cell = new Cell();
                            cell.StyleIndex = digitStyle;
                            cell.DataType = CellValues.Number;
                            if (percentFractionalCell.Value.HasValue)
                            {
                                cell.CellValue = new CellValue(DecimalValue.FromDecimal(percentFractionalCell.Value.Value));
                            }


                            cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                            row.Append(cell);

                            continue;
                        }
                        #endregion

                        #region date
                        dateCell = excelCell as ExcelCellDate;
                        if (dateCell != null)
                        {
                            cell = new Cell();
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            cell.StyleIndex = dateStyle;

                            if (dateCell.Value.HasValue)
                            {
                                cell.CellValue = new CellValue(DoubleValue.FromDouble(dateCell.Value.Value.ToOADate()));
                            }

                            cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                            row.Append(cell);
                            continue;
                        }
                        #endregion

                        #region date an time
                        dateAndTimeCell = excelCell as ExcelCellDateAndTime;
                        if (dateAndTimeCell != null)
                        {
                            cell = new Cell();
                            cell.StyleIndex = dateStyle;
                            cell.DataType = null;
                            if (dateAndTimeCell.Value.HasValue)
                            {
                                cell.CellValue = new CellValue(DoubleValue.FromDouble(dateAndTimeCell.Value.Value.ToOADate()));
                            }


                            cell.CellReference = ExcelOpenXmlSaxReader.GetCellReference(currentRowIndex, currentCellIndex++);
                            row.Append(cell);

                            continue;
                        }
                        #endregion
                    }
                }
                worksheetPart.Worksheet.Save();
            }
        }
    }
}
