using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HelperLibrary.ExcelOpenXml
{
    /// <summary>
    /// Decimal separator for numbers.
    /// </summary>
    public enum ExcelOpenXmlDecimalSeparator
    {
        /// <summary>
        /// Numbers are delimited with dot (period).
        /// </summary>
        Dot,

        /// <summary>
        /// Numbers are delimited with comma.
        /// </summary>
        Comma
    }

    /// <summary>
    /// Writes data to Excel workbooks via OpenXML (uses SAX, supports large files).
    /// </summary>
    public class ExcelOpenXmlSaxReader : IDisposable
    {
        private static System.Globalization.NumberFormatInfo dotNumberFormatInfo;
        private static System.Globalization.NumberFormatInfo commaNumberFormatInfo;

        private DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheetDocument;
        private DocumentFormat.OpenXml.OpenXmlReader reader;
        private System.IO.FileStream fileStream;
        private string fileName;

        private int columnCount, rowCount;

        private DocumentFormat.OpenXml.Spreadsheet.Cell currentCell;
        private int currentCellIndex;
        private int currentCellPosition;

        private int currentRowIndex;
        private int currentRowPosition;

        private string[] sharedStrings;

        static ExcelOpenXmlSaxReader()
        {
            dotNumberFormatInfo = new System.Globalization.NumberFormatInfo();
            dotNumberFormatInfo.NumberDecimalSeparator = ".";
            dotNumberFormatInfo.NumberGroupSeparator = string.Empty;

            commaNumberFormatInfo = new System.Globalization.NumberFormatInfo();
            commaNumberFormatInfo.NumberDecimalSeparator = ",";
            commaNumberFormatInfo.NumberGroupSeparator = string.Empty;
        }

        /// <summary>
        /// File name of current workbook.
        /// </summary>
        public string FileName
        {
            get
            {
                return fileName;
            }
        }

        /// <summary>
        /// Number of columns in current worksheet.
        /// </summary>
        public int ColumnCount
        {
            get
            {
                return columnCount;
            }
        }

        /// <summary>
        /// Number of rows in current worksheet.
        /// </summary>
        public int RowCount
        {
            get
            {
                return rowCount;
            }
        }

        /// <summary>
        /// Current column in current worksheet.
        /// </summary>
        public int CurrentColumn
        {
            get
            {
                return currentCellPosition;
            }
        }

        /// <summary>
        /// Current row in current worksheet.
        /// </summary>
        public int CurrentRow
        {
            get
            {
                return currentRowIndex;
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
        /// Opens workbook in read-only mode.
        /// <remarks>
        /// Can open files that are already opened in other programs.
        /// </remarks>
        /// </summary>
        /// <param name="fileName">Workbook file name.</param>
        /// <returns>True if success.</returns>
        public bool OpenWorkbook(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException("fileName", "File name cannot be null");
            }

            CleanupResources();

            this.fileName = fileName;

            bool result = false;

            try
            {
                fileStream = new System.IO.FileStream(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);

                spreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(fileStream, false);
                //this.spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                //this.spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;

                result = spreadsheetDocument != null;
            }
            finally
            {
                if (!result)
                {
                    CleanupResources();
                }
            }

            return result;
        }

        /// <summary>
        /// Opens worksheet by index.
        /// </summary>
        /// <param name="index">Worksheet index (one-based).</param>
        /// <returns>True if success.</returns>
        public bool OpenWorksheet(int index)
        {
            if (spreadsheetDocument == null)
            {
                return false;
            }

            if (index < 1 || index > spreadsheetDocument.WorkbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count())
            {
                throw new ArgumentOutOfRangeException("index", "Worksheet index is out of range");
            }

            columnCount = 0;
            rowCount = 0;

            currentCellIndex = 0;
            currentCellPosition = 0;
            currentRowIndex = 0;
            currentRowPosition = 0;

            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ElementAt(index - 1);

            foreach (var worksheetPart in spreadsheetDocument.WorkbookPart.WorksheetParts)
            {
                if (spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart) == sheet.Id)
                {
                    reader = DocumentFormat.OpenXml.OpenXmlReader.Create(worksheetPart);

                    if (reader == null)
                    {
                        return false;
                    }

                    if (!ReadSheetDimensions())
                    {
                        return false;
                    }

                    CacheSharedStrings();

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Opens worksheet by name.
        /// </summary>
        /// <param name="index">Worksheet name.</param>
        /// <returns>True if success.</returns>
        public bool OpenWorksheet(string name, bool isIgnorWorksheet = false)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentOutOfRangeException("name", string.Format("Worksheet name should not be empty in {0}", fileName));
            }

            if (spreadsheetDocument == null)
            {
                return false;
            }

            columnCount = 0;
            rowCount = 0;

            currentCellIndex = 0;
            currentCellPosition = 0;
            currentRowIndex = 0;
            currentRowPosition = 0;

            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault(s => string.Compare(s.Name, name, true) == 0);
            if (isIgnorWorksheet)
            {
                if (sheet == null)
                {
                    return false;
                }
            }
            else
            {
                if (sheet == null)
                {
                    throw new ArgumentOutOfRangeException("name", string.Format("Worksheet {0} name was not found in {1}", name, fileName));
                }
            }

            foreach (var worksheetPart in spreadsheetDocument.WorkbookPart.WorksheetParts)
            {
                if (spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart) == sheet.Id)
                {
                    reader = DocumentFormat.OpenXml.OpenXmlReader.Create(worksheetPart);

                    if (reader == null)
                    {
                        return false;
                    }

                    if (!ReadSheetDimensions())
                    {
                        return false;
                    }

                    CacheSharedStrings();

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Moves to the next row of current worksheet.
        /// </summary>
        /// <returns>True if success.</returns>
        public bool MoveToNextRow(bool allowEmptyRow = false)
        {
            currentCellPosition = 1;

            if (currentRowPosition < currentRowIndex)
            {
                // This row does not exist in worksheet (the row is empty)
                currentRowPosition++;
                return true;
            }

            Type rowType = typeof(DocumentFormat.OpenXml.Spreadsheet.Row);
            Type cellType = typeof(DocumentFormat.OpenXml.Spreadsheet.Cell);

            while (reader.Read())
            {
                if (reader.IsStartElement && (reader.ElementType == rowType))
                {
                    // Read next available row
                    DocumentFormat.OpenXml.OpenXmlAttribute attribute = reader.Attributes
                        .First(a => a.LocalName == "r");

                    currentRowIndex = Convert.ToInt32(attribute.Value);

                    // Fix: for first row with zero index (invalid) 
                    if (currentRowIndex == 0)
                    {
                        currentRowIndex = 1;
                    }

                    currentRowPosition++;

                    currentCell = null;
                    currentCellIndex = 0;
                    currentCellPosition = 0;

                    if (reader.ReadFirstChild())
                    {
                        // Row contains cells
                        if (reader.IsStartElement && (reader.ElementType == cellType))
                        {
                            // Read first available cell
                            currentCell = (DocumentFormat.OpenXml.Spreadsheet.Cell)reader.LoadCurrentElement();
                            currentCellIndex = ColumnNameToIndex(GetColumnNameFromCellReference(currentCell.CellReference));
                            currentCellPosition = 1;

                            // Move to the next cell
                            reader.ReadNextSibling();

                            return true;
                        }
                    }
                    else
                    {
                        // Row is empty
                        return allowEmptyRow;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Moves to the specified row (relative to current row) of current worksheet.
        /// </summary>
        /// <returns>True if success.</returns>
        public bool MoveToNextRow(int count)
        {
            if (count <= 0)
            {
                throw new ArgumentOutOfRangeException("count", "Row count cannot be zero or negative value");
            }

            for (int i = 0; i < count; i++)
            {
                if (!MoveToNextRow(true))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Reads string value from the next cell.
        /// </summary>
        /// <returns>Null value if cell is empty.</returns>
        public string ReadCellStringValue()
        {
            string result = null;

            if (currentCell != null)
            {
                if (currentCellIndex == currentCellPosition)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell c = currentCell;

                    if (c.DataType != null && c.DataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
                    {
                        // Shared string (stored in separate table)
                        ////DocumentFormat.OpenXml.Spreadsheet.SharedStringItem ssi = this.spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>()
                        ////    .ElementAt(int.Parse(c.CellValue.InnerText));

                        ////result = ssi.InnerText;

                        result = sharedStrings[int.Parse(c.CellValue.InnerText)];
                    }
                    else
                    {
                        // Inline string (stored as child node)
                        if (c.DataType != null && c.DataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString)
                        {
                            result = c.FirstChild.InnerText;
                        }
                        else
                        {
                            // Regular string
                            if (c.CellValue != null)
                            {
                                result = c.CellValue.InnerText;
                            }
                        }
                    }
                }
            }

            if (!SkipCells(1))
            {
                currentCell = null;
                currentCellIndex = 0;
            }

            return result;
        }

        /// <summary>
        /// Reads date and time value from the next cell.
        /// </summary>
        /// <returns>Null value if cell is empty.</returns>
        public DateTime? ReadCellDateValue(bool allowNotaDate = false)
        {
            DateTime? result = null;

            if (currentCell != null)
            {
                if (currentCellIndex == currentCellPosition)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell c = currentCell;

                    string s = GetCurrentCellAsString();
                    if (!string.IsNullOrEmpty(s)) //!= null
                    {
                        if (allowNotaDate)
                        {
                            result = null;
                        }
                        else
                        {
                            throw new FormatException(string.Format(
                                "Cell value \"{0}\" at {1} is not a date in \"{2}\"",
                                s,
                                GetCellReference(currentRowPosition, currentCellPosition),
                                fileName));
                        }
                    }

                    if (c.CellValue != null)
                    {
                        try
                        {
                            result = DateTime.FromOADate(Convert.ToDouble(c.CellValue.InnerText, dotNumberFormatInfo));
                        }
                        catch
                        {
                            if (allowNotaDate)
                            {
                                result = null;
                            }
                            else
                            {
                                throw new FormatException(string.Format(
                                    "Cell value \"{0}\" at {1} is not a date in \"{2}\"",
                                    s,
                                    GetCellReference(currentRowPosition, currentCellPosition),
                                    fileName));
                            }
                        }
                    }
                    else
                    {
                        result = null;
                    }
                }
            }

            if (!SkipCells(1))
            {
                currentCell = null;
                currentCellIndex = 0;
            }

            return result;
        }

        /// <summary>
        /// Reads integral number from the next cell.
        /// </summary>
        /// <returns>Null value if cell is empty.</returns>        
        public int? ReadCellIntegralNumberValue()
        {
            int? result = null;

            if (currentCell != null)
            {
                if (currentCellIndex == currentCellPosition)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell c = currentCell;

                    string s = GetCurrentCellAsString();
                    if (s == null)
                    {
                        if (c.CellValue != null)
                        {
                            s = c.CellValue.InnerText;
                        }
                    }

                    if (s != null)
                    {
                        try
                        {
                            result = Convert.ToInt32(s, dotNumberFormatInfo);
                        }
                        catch
                        {
                            throw new FormatException(string.Format(
                                "Cell value \"{0}\" at {1} is not an integral number in \"{2}\"",
                                s,
                                GetCellReference(currentRowPosition, currentCellPosition),
                                fileName));
                        }
                    }
                }
            }

            if (!SkipCells(1))
            {
                currentCell = null;
                currentCellIndex = 0;
            }

            return result;
        }

        /// <summary>
        /// Reads fractional number from the next cell.
        /// </summary>
        /// <returns>Null value if cell is empty.</returns>        
        public double? ReadCellFractionalNumberValue(ExcelOpenXmlDecimalSeparator separator, bool convertSeparator)
        {
            double? result = null;

            if (currentCell != null)
            {
                if (currentCellIndex == currentCellPosition)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell c = currentCell;

                    string s = GetCurrentCellAsString();
                    if (s == null)
                    {
                        if (c.CellValue != null)
                        {
                            s = c.CellValue.InnerText;
                        }
                    }

                    if (s != null)
                    {
                        if (convertSeparator)
                        {
                            switch (separator)
                            {
                                case ExcelOpenXmlDecimalSeparator.Dot:
                                    s = s.Replace(',', '.');
                                    break;
                                case ExcelOpenXmlDecimalSeparator.Comma:
                                    s = s.Replace('.', ',');
                                    break;
                            }
                        }

                        try
                        {
                            result = double.Parse(s, separator == ExcelOpenXmlDecimalSeparator.Dot ? dotNumberFormatInfo : commaNumberFormatInfo);
                            ////result = Convert.ToDouble(s, separator == ExcelOpenXmlDecimalSeparator.Dot ? dotNumberFormatInfo : commaNumberFormatInfo);
                        }
                        catch
                        {
                            throw new FormatException(string.Format(
                                "Cell value \"{0}\" at {1} is not a fractional number in \"{2}\"",
                                s,
                                GetCellReference(currentRowPosition, currentCellPosition),
                                fileName));
                        }
                    }
                }
            }

            if (!SkipCells(1))
            {
                currentCell = null;
                currentCellIndex = 0;
            }

            return result;
        }

        /// <summary>
        /// Reads fractional number from the next cell.
        /// </summary>
        /// <returns>Null value if cell is empty.</returns>        
        public double? ReadCellFractionalNumberValue()
        {
            return ReadCellFractionalNumberValue(ExcelOpenXmlDecimalSeparator.Dot, false);
        }

        /// <summary>
        /// Reads decimal number from the next cell.
        /// </summary>
        /// <returns>Null value if cell is empty.</returns>        
        public decimal? ReadCellDecimalNumberValue(ExcelOpenXmlDecimalSeparator separator, bool convertSeparator)
        {
            decimal? result = null;

            if (currentCell != null)
            {
                if (currentCellIndex == currentCellPosition)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell c = currentCell;

                    string s = GetCurrentCellAsString();
                    if (s == null)
                    {
                        if (c.CellValue != null)
                        {
                            s = c.CellValue.InnerText;
                        }
                    }

                    if (s != null)
                    {
                        if (convertSeparator)
                        {
                            switch (separator)
                            {
                                case ExcelOpenXmlDecimalSeparator.Dot:
                                    s = s.Replace(',', '.');
                                    break;
                                case ExcelOpenXmlDecimalSeparator.Comma:
                                    s = s.Replace('.', ',');
                                    break;
                            }
                        }

                        try
                        {
                            result = decimal.Parse(s, System.Globalization.NumberStyles.Float, separator == ExcelOpenXmlDecimalSeparator.Dot ? dotNumberFormatInfo : commaNumberFormatInfo);
                            ////result = Convert.ToDecimal(s, separator == ExcelOpenXmlDecimalSeparator.Dot ? dotNumberFormatInfo : commaNumberFormatInfo);
                        }
                        catch
                        {
                            throw new FormatException(string.Format(
                                "Cell value \"{0}\" at {1} is not a decimal number in \"{2}\"",
                                s,
                                GetCellReference(currentRowPosition, currentCellPosition),
                                fileName));
                        }
                    }
                }
            }

            if (!SkipCells(1))
            {
                currentCell = null;
                currentCellIndex = 0;
            }

            return result;
        }

        /// <summary>
        /// Reads decimal number from the next cell.
        /// </summary>
        /// <returns>Null value if cell is empty.</returns>        
        public decimal? ReadCellDecimalNumberValue()
        {
            return ReadCellDecimalNumberValue(ExcelOpenXmlDecimalSeparator.Dot, false);
        }

        /// <summary>
        /// Skips specified number of cells.
        /// </summary>
        /// <param name="count">Number of cells to skip.</param>
        /// <returns>True if success.</returns>
        public bool SkipCells(int count)
        {
            if (count <= 0)
            {
                throw new ArgumentOutOfRangeException("count", "Cell count cannot be zero or negative value");
            }

            for (int i = 0; i < count; i++)
            {
                if (currentCellPosition < currentCellIndex)
                {
                    // This cell does not exist in worksheet (the cell is empty)
                    currentCellPosition++;
                    continue;
                }

                if (!reader.EOF && reader.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Cell) && reader.IsStartElement)
                {
                    // Read next available cell
                    currentCell = (DocumentFormat.OpenXml.Spreadsheet.Cell)reader.LoadCurrentElement();
                    currentCellIndex = ColumnNameToIndex(GetColumnNameFromCellReference(currentCell.CellReference));
                    currentCellPosition++;

                    // Move to the next cell
                    reader.ReadNextSibling();
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Reads column indices from the current row.
        /// </summary>
        /// <param name="columnDefinitions">List of column definitions.</param>
        /// <param name="missmatchColumn">True/False if columns is missing(like 6 columns irl but 0 by defenition).</param>
        /// <param name="missmatchColumnCount">Count of columns if is missmatch</param>
        /// <returns>True if all column indices are found, False otherwise.</returns>
        public bool ReadColumnDefinitions(List<ExcelColumnDefinition> columnDefinitions, bool missmatchColumn = false, int missmatchColumnCount = -1)
        {
            ExcelColumnDefinition columnDefinition;
            string fieldName;
            int count = 0;
            int curentColl;
            while (ColumnCount > count || (missmatchColumnCount > count && missmatchColumn))
            {
                curentColl = CurrentColumn;
                fieldName = ReadCellStringValue();
                fieldName = !string.IsNullOrEmpty(fieldName) ? fieldName.Replace("\n", "") : "";
                fieldName = fieldName.Trim();

                columnDefinition = columnDefinitions.Where(q => !q.IsManyVariant)
                    .FirstOrDefault(cd => string.Compare(cd.Caption, fieldName, true) == 0 && cd.Index == 0);
                if (columnDefinitions.Where(q => q.IsManyVariant).Count() > 0 && columnDefinition == null)
                {
                    columnDefinition = columnDefinitions.Where(q => q.IsManyVariant)
                        .FirstOrDefault(cd => cd.Variants.Contains(fieldName) && cd.Index == 0);
                }


                if (columnDefinition != null && columnDefinition.Skip == 0)
                {
                    columnDefinition.Index = curentColl;
                }
                else if (columnDefinition != null && columnDefinition.Skip != 0)
                {
                    columnDefinition.Skip--;
                }
                count++;
                if (columnDefinitions.Where(q => q.Index == 0).Count() == 0)
                {
                    break;
                }
            }
            bool result = columnDefinitions.Count(cd => cd.Index == 0) == 0;
            if (!result)
            {
                string outputString = "В файле {0} отсутствуют необходимые колонки!" + Environment.NewLine + "Нужны столбцы:";
                string errorString = "";
                foreach (var column in columnDefinitions.Where(q => q.Index == 0))
                {
                    outputString += Environment.NewLine + column.Caption;
                    errorString += column.Caption + Environment.NewLine;
                }
                ReadColumndDefenitionResult = outputString;
                ReadColumnDefenitionColumnError = errorString;
            }
            return result;
        }

        private bool ReadColumnDefinitions(List<ExcelColumnDefinition> columnDefinitions, int depth)
        {
            ExcelColumnDefinition columnDefinition;
            string fieldName;
            int curentDepth = 0;
            bool isDone = false;
            do
            {
                int count = 0;
                if (curentDepth != 0)
                {
                    MoveToNextRow();
                }

                while (ColumnCount > count)
                {
                    fieldName = ReadCellStringValue();
                    fieldName = !string.IsNullOrEmpty(fieldName) ? fieldName.Replace("\n", "") : fieldName;
                    columnDefinition = columnDefinitions
                        .FirstOrDefault(cd => string.Compare(cd.Caption, fieldName, true) == 0);
                    if (columnDefinition != null)
                    {
                        columnDefinition.Index = CurrentColumn - 1;
                        if (ColumnCount == count + 1)
                        {
                            columnDefinition.Index++;
                        }
                    }
                    count++;
                    if (columnDefinitions.Where(q => q.Index == 0).Count() == 0)
                    {
                        isDone = true;
                        break;
                    }
                }
                curentDepth++;
            } while (curentDepth < depth || !isDone);


            bool result = columnDefinitions.Count(cd => cd.Index == 0) == 0;
            if (!result)
            {
                string outputString = "В файле {0} отсутствуют необходимые колонки!" + Environment.NewLine + "Нужны столбцы:";
                string errorString = "";
                foreach (var column in columnDefinitions.Where(q => q.Index == 0))
                {
                    outputString += Environment.NewLine + column.Caption;
                    errorString += column.Caption + Environment.NewLine;
                }
                ReadColumndDefenitionResult = outputString;
                ReadColumnDefenitionColumnError = errorString;
            }

            return result;
        }


        public string ReadColumndDefenitionResult { get; set; }
        public string ReadColumnDefenitionColumnError { get; set; }

        /// <summary>
        /// Reads column values from the current row.
        /// </summary>
        /// <param name="columnDefinitions">List of column definitions.</param>
        /// <returns>True if all column values are found, False otherwise.</returns>
        public bool ReadColumns(List<ExcelColumnDefinition> columnDefinitions, bool ignoreDate = false, bool missmatchColumn = false, int missmatchColumnCount = -1)
        {
            foreach (ExcelColumnDefinition columnDefinition2 in columnDefinitions)
            {
                if (columnDefinition2.CellValue is ExcelCellString)
                {
                    columnDefinition2.CellValue = new ExcelCellString(null);
                }

                if (columnDefinition2.CellValue is ExcelCellNumberIntegral)
                {
                    columnDefinition2.CellValue = new ExcelCellNumberIntegral(null);
                }

                if (columnDefinition2.CellValue is ExcelCellNumberFractional)
                {
                    columnDefinition2.CellValue = new ExcelCellNumberFractional(null);
                }

                if (columnDefinition2.CellValue is ExcelCellPercentIntegral)
                {
                    columnDefinition2.CellValue = new ExcelCellPercentIntegral(null);
                }

                if (columnDefinition2.CellValue is ExcelCellPercentFractional)
                {
                    columnDefinition2.CellValue = new ExcelCellPercentFractional(null);
                }

                if (columnDefinition2.CellValue is ExcelCellDate)
                {
                    columnDefinition2.CellValue = new ExcelCellDate(null);
                }

                if (columnDefinition2.CellValue is ExcelCellDateAndTime)
                {
                    columnDefinition2.CellValue = new ExcelCellDateAndTime(null);
                }
            }

            ExcelColumnDefinition columnDefinition;
            int columncount = missmatchColumn ? missmatchColumnCount : ColumnCount;

            for (int i = 1; i <= columncount; i++)
            {
                columnDefinition = columnDefinitions
                    .FirstOrDefault(cd => cd.Index == i);
                if (columnDefinition != null)
                {
                    if (columnDefinition.CellValue is ExcelCellString)
                    {
                        columnDefinition.CellValue = new ExcelCellString(ReadCellStringValue());
                        continue;
                    }

                    if (columnDefinition.CellValue is ExcelCellNumberIntegral)
                    {
                        columnDefinition.CellValue = new ExcelCellNumberIntegral(ReadCellIntegralNumberValue());
                        continue;
                    }

                    if (columnDefinition.CellValue is ExcelCellNumberFractional)
                    {
                        columnDefinition.CellValue = new ExcelCellNumberFractional(ReadCellDecimalNumberValue());
                        continue;
                    }

                    if (columnDefinition.CellValue is ExcelCellPercentIntegral)
                    {
                        columnDefinition.CellValue = new ExcelCellPercentIntegral(ReadCellIntegralNumberValue());
                        continue;
                    }

                    if (columnDefinition.CellValue is ExcelCellPercentFractional)
                    {
                        columnDefinition.CellValue = new ExcelCellPercentFractional(ReadCellDecimalNumberValue());
                        continue;
                    }

                    if (columnDefinition.CellValue is ExcelCellDate)
                    {
                        columnDefinition.CellValue = new ExcelCellDate(ReadCellDateValue(ignoreDate));
                        continue;
                    }

                    if (columnDefinition.CellValue is ExcelCellDateAndTime)
                    {
                        columnDefinition.CellValue = new ExcelCellDateAndTime(ReadCellDateValue());
                        continue;
                    }
                }

                if (!SkipCells(1))
                {
                    break;
                }
            }

            return columnDefinitions
                .Count(cd => cd.CellValue == null) == 0;
        }


        /// <summary>
        /// Converts column index to column name.
        /// </summary>
        /// <param name="index">Column index (e.g. "1").</param>
        /// <returns>Column name (e.g. "A").</returns>
        internal static string ColumnIndexToName(int index)
        {
            if (index < 1 || index > 16384)
            {
                throw new ArgumentOutOfRangeException("count", "Column index is out of range (1...16384)");
            }

            int dividend = index;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        /// <summary>
        /// Converts column name to column index.
        /// </summary>
        /// <param name="name">Column name (e.g. "A").</param>
        /// <returns>Column index (e.g. "1").</returns>
        internal static int ColumnNameToIndex(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentNullException("name");
            }

            name = name.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < name.Length; i++)
            {
                sum *= 26;
                sum += name[i] - 'A' + 1;
            }

            if (sum < 1 || sum > 16384)
            {
                throw new ArgumentOutOfRangeException("count", "Result is out of range (1...16384)");
            }

            return sum;
        }

        /// <summary>
        /// Returns cell reference string.
        /// </summary>
        /// <param name="rowIndex">Row index (one-based).</param>
        /// <param name="columnIndex">Column index (one-based).</param>
        /// <returns>Cell reference string (e.g. "A1").</returns>
        internal static string GetCellReference(int rowIndex, int columnIndex)
        {
            return ColumnIndexToName(columnIndex) + rowIndex.ToString();
        }

        /// <summary>
        /// Returns column name part from cell reference.
        /// </summary>
        /// <param name="cellReference">Cell reference string (e.g. "A1").</param>
        /// <returns>Column name (e.g. "A").</returns>
        internal static string GetColumnNameFromCellReference(string cellReference)
        {
            StringBuilder result = new StringBuilder();

            cellReference = cellReference.ToUpperInvariant();

            char c;
            for (int i = 0; i < cellReference.Length; i++)
            {
                c = cellReference[i];

                if ((c < 'A') || (c > 'Z'))
                {
                    break;
                }
                else
                {
                    result.Append(c);
                }
            }

            return result.ToString();
        }

        /// <summary>
        /// Returns row name part from cell reference.
        /// </summary>
        /// <param name="cellReference">Cell reference string (e.g. "A1").</param>
        /// <returns>Row name (e.g. "1").</returns>
        internal static string GetRowNameFromCellReference(string cellReference)
        {
            StringBuilder result = new StringBuilder();

            cellReference = cellReference.ToUpperInvariant();

            char c;
            for (int i = 0; i < cellReference.Length; i++)
            {
                c = cellReference[i];

                if ((c >= '0') && (c <= '9'))
                {
                    result.Append(c);
                }
            }

            return result.ToString();
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
            fileName = null;

            currentCell = null;

            if (reader != null)
            {
                reader.Close();
                reader.Dispose();
                reader = null;
            }

            if (spreadsheetDocument != null)
            {
                spreadsheetDocument.Close();
                spreadsheetDocument.Dispose();
                spreadsheetDocument = null;
            }

            if (fileStream != null)
            {
                fileStream.Close();
                fileStream.Dispose();
                fileStream = null;
            }

            sharedStrings = null;
        }

        private bool ReadSheetDimensions()
        {
            while (reader.Read())
            {
                if (reader.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.SheetDimension))
                {
                    // Read sheet dimensions
                    DocumentFormat.OpenXml.OpenXmlAttribute attribute = reader.Attributes
                        .First(a => a.LocalName == "ref");

                    string[] words = attribute.Value.Split(new char[] { ':' });

                    if (words.Length < 2)
                    {
                        columnCount = 0;
                        rowCount = 0;

                        return true;
                    }

                    columnCount = ColumnNameToIndex(GetColumnNameFromCellReference(words[1]));
                    rowCount = Convert.ToInt32(GetRowNameFromCellReference(words[1]));

                    /*
                    int fromColumn = ColumnNameToIndex(GetColumnNameFromCellReference(words[0]));
                    int toColumn = ColumnNameToIndex(GetColumnNameFromCellReference(words[1]));

                    this.columnCount = toColumn - fromColumn + 1;

                    int fromRow = Convert.ToInt32(GetRowNameFromCellReference(words[0]));
                    int toRow = Convert.ToInt32(GetRowNameFromCellReference(words[1]));

                    this.rowCount = toRow - fromRow + 1;
                    */

                    return true;
                }

                if (reader.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.SheetData))
                {
                    // Sheet data reached, stop searching for sheet dimensions
                    return true;
                }
            }

            return false;
        }

        private string GetCurrentCellAsString()
        {
            DocumentFormat.OpenXml.Spreadsheet.Cell c = currentCell;

            if (c.DataType != null && c.DataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
            {
                // Shared string (stored in separate table)
                DocumentFormat.OpenXml.Spreadsheet.SharedStringItem ssi = spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>()
                    .ElementAt(int.Parse(c.CellValue.InnerText));

                return ssi.InnerText;
            }

            if (c.DataType != null && c.DataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString)
            {
                // Inline string (stored as child node)
                return c.FirstChild.InnerText;
            }

            return null;
        }

        private void CacheSharedStrings()
        {
            sharedStrings = null;

            if (spreadsheetDocument.WorkbookPart.SharedStringTablePart != null)
            {
                sharedStrings = new string[spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements().Count()];

                int i = 0;
                foreach (var element in spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements())
                {
                    sharedStrings[i++] = element.InnerText;
                }
            }
        }
    }
}
