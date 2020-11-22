using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace SpreadsheetWrapper
{
    /// <summary>
    /// Build and export spreadsheets to Excel, CSV or Byte Array
    /// </summary>
    public class Spreadsheet : IDisposable
    {
        public ExcelPackage Workbook { get; private set; }

        public Spreadsheet()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Workbook = new ExcelPackage();
        }        

        /// <summary>
        /// Returns the specified sheet as byte representation of the equivalent Excel file
        /// </summary>
        /// <returns></returns>
        public byte[] SerializeExcel()
        {
            if(Workbook.Workbook.Worksheets.Count == 0)
            {
                _ = GetSheetByName("No data");
            }
            return Workbook.GetAsByteArray();
        }

        /// <summary>
        /// Returns the specified sheet as byte representation of the equivalent CSV UTF-8 encoded text content
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public byte[] SerializeCsv(string sheetName, char separator = ';')
        {
            var sheet = GetSheetByName(sheetName);
            return SerializeCsv(sheet, separator);
        }

        /// <summary>
        /// Returns the specified sheet as byte representation of the equivalent CSV UTF-8 encoded text content
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public byte[] SerializeCsv(ExcelWorksheet sheet, char separator = ';')
        {
            var content = GetSheetAsCsv(sheet, separator);
            return content.SelectMany(x => Encoding.UTF8.GetBytes(x)).ToArray();
        }

        /// <summary>
        /// Save the document as Excel file
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="basePath"></param>
        /// <param name="password"></param>
        public void SaveExcelAs(string fileName, string basePath = @".\", string password = "")
        {
            fileName = GetValidFileName(fileName);
            if(!fileName.ToLower().EndsWith(".xlsx"))
            {
                fileName += ".xlsx";
            }
            Directory.CreateDirectory(basePath);
            string filePath = Path.Combine(basePath, fileName);
            var fInfo = new FileInfo(filePath);

            if (string.IsNullOrWhiteSpace(password))
            {
                Workbook.SaveAs(fInfo);
            }
            else
            {
                Workbook.SaveAs(fInfo, password);
            }
        }

        /// <summary>
        /// Saves the document in one (or more) files, one file per sheet
        /// </summary>        
        /// <param name="basePath"></param>
        /// <param name="separator"></param>
        /// <param name="append"></param>
        public void SaveCsvAs(string basePath = @".\", char separator = ';', bool append = false)
        {
            var sheetNames = GetSheetNames();
            foreach (string sheetName in sheetNames)
            {
                SaveCsvAs(GetSheetByName(sheetName), sheetName, basePath, separator, append);
            }
        }

        /// <summary>
        /// Save the specified sheet as CSV text file
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="fileName"></param>
        /// <param name="basePath"></param>
        /// <param name="separator"></param>
        /// <param name="append"></param>
        public void SaveCsvAs(string sheetName, string fileName, string basePath = @".\", char separator = ';', bool append = false)
        {
            var sheet = GetSheetByName(sheetName);
            SaveCsvAs(sheet, fileName, basePath, separator, append);
        }

        /// <summary>
        /// Save the specified sheet as CSV UTF-8 encoded text file
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fileName"></param>
        /// <param name="basePath"></param>
        /// <param name="separator"></param>
        /// <param name="append"></param>
        public void SaveCsvAs(ExcelWorksheet sheet, string fileName, string basePath = @".\", char separator = ';', bool append = false)
        {
            fileName = GetValidFileName(fileName);
            if (!fileName.ToLower().EndsWith(".csv"))
            {
                fileName += ".csv";
            }
            Directory.CreateDirectory(basePath);
            string filePath = Path.Combine(basePath, fileName);
            string content = GetSheetAsCsvString(sheet, separator) + Environment.NewLine;
            if (!append)
            {
                File.WriteAllText(filePath, content, Encoding.UTF8);
            }
            else
            {
                File.AppendAllText(filePath, content, Encoding.UTF8);
            }
        }

        /// <summary>
        /// Returns the specified sheet as collection of CSV formatted text
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public string GetSheetAsCsvString(string sheetName, char separator = ';')
        {
            var sheet = GetSheetByName(sheetName);
            return GetSheetAsCsvString(sheet, separator);
        }

        /// <summary>
        /// Returns the specified sheet as collection of CSV formatted text
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public string GetSheetAsCsvString(ExcelWorksheet sheet, char separator = ';')
        {
            var csv = GetSheetAsCsv(sheet, separator);
            var content = string.Join(Environment.NewLine, csv);
            return content;
        }

        /// <summary>
        /// Returns the specified sheet as collection of CSV formatted strings
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public List<string> GetSheetAsCsv(string sheetName, char separator = ';')
        {
            var sheet = GetSheetByName(sheetName);
            return GetSheetAsCsv(sheet, separator);
        }

        /// <summary>
        /// Returns the specified sheet as collection of CSV formatted strings
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public List<string> GetSheetAsCsv(ExcelWorksheet sheet, char separator = ';')
        {
            var content = new List<string>();
            int totalLines = sheet.Dimension.End.Row;

            if (totalLines == 0)
            {
                content.Add("No data");
            }
            else
            {
                for (int i = 1; i <= totalLines; i++)
                {
                    var row = sheet.Cells[string.Format("{0}:{0}", i)];
                    var csvLineValues = row.Select(x => $"\"{x.Text.Replace("\"", "\"\"")}\"");
                    var csvLineText = string.Join(separator, csvLineValues);
                    content.Add(csvLineText);
                }
            }
            return content;
        }

        /// <summary>
        /// Will insert bulk data in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="table">DataTable object to dump</param>
        /// <param name="includeHeader">Generate header using DataTable ColumnName attribute</param>
        /// <param name="rowConverter">Function with argument object[] and returns string[]. Specify when default does not yield desired format</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="autofit">Should adjust all content columns width</param>
        public void InsertTable(DataTable table, bool includeHeader = true, Func<object[], string[]> rowConverter = null, int rowIndex = 2, int columnIndex = 1, bool autofit = false)
        {
            InsertTable(table, table.TableName, includeHeader, rowConverter, rowIndex, columnIndex, autofit);
        }

        /// <summary>
        /// Will insert bulk data in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="table">DataTable object to dump</param>
        /// <param name="sheetName">Sheet to insert data. If the sheet is not found, it will be created</param>
        /// <param name="includeHeader">Generate header using DataTable ColumnName attribute</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="rowConverter">Function with argument object[] and returns string[]. Specify when default does not yield desired format</param>
        /// <param name="autofit">Should adjust all content columns width</param>
        public void InsertTable(DataTable table, string sheetName, bool includeHeader = true, Func<object[], string[]> rowConverter = null, int rowIndex = 2, int columnIndex = 1, bool autofit = false)
        {
            ExcelWorksheet sheet;

            if (includeHeader)
            {
                int headerIndex = rowIndex - 1;
                if (headerIndex == 0)
                {
                    headerIndex = rowIndex;
                    rowIndex++;
                }

                var header = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    header.Add(column.ColumnName);
                }

                sheet = GetSheetByNameWithHeader(sheetName, header.ToArray(), headerIndex, columnIndex);
                
            }
            else
            {
                sheet = GetSheetByName(sheetName);
            }

            InsertTable(table, sheet, rowConverter, rowIndex, columnIndex, autofit);
        }

        /// <summary>
        /// Will insert bulk data in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="table">DataTable object to dump</param>
        /// <param name="sheet">Sheet to insert data</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="rowConverter">Function with argument object[] and returns string[]. Specify when default does not yield desired format. Specify when default does not yield desired format</param>
        /// <param name="autofit">Should adjust all content columns width</param>
        public void InsertTable(DataTable table, ExcelWorksheet sheet, Func<object[], string[]> rowConverter = null, int rowIndex = 2, int columnIndex = 1, bool autofit = false)
        {
            if (rowIndex <= 0) rowIndex = 1;
            if (columnIndex <= 0) columnIndex = 1;

            string range = GetRangeForCell(rowIndex, columnIndex);
            if (rowConverter == null)
            {
                sheet.Cells[range].LoadFromDataTable(table);
                if (autofit)
                {
                    Autofit(sheet);
                }
            }
            else
            {
                var content = new List<string[]>();
                foreach (DataRow row in table.Rows)
                {
                    content.Add(rowConverter.Invoke(row.ItemArray));
                }
                InsertRows(sheet, content, rowIndex, columnIndex, autofit);
            }            
        }

        /// <summary>
        /// Will insert data parameter in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="sheetName">Sheet to insert data. If the sheet is not found, it will be created</param>
        /// <param name="data">Data to insert in sheet</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="autofit">Should adjust all content columns width</param>
        public void InsertRow(string sheetName, string[] data, int rowIndex = 2, int columnIndex = 1, bool autofit = false)
        {
            InsertRows(sheetName, new List<string[]> { data }, rowIndex, columnIndex, autofit);
        }

        /// <summary>
        /// Will insert data parameter in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="sheet">Sheet to insert data</param>
        /// <param name="data">Data to insert in sheet</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="autofit">Should adjust all content columns width</param>
        public void InsertRow(ExcelWorksheet sheet, string[] data, int rowIndex = 2, int columnIndex = 1, bool autofit = false)
        {
            InsertRows(sheet, new List<string[]> { data }, rowIndex, columnIndex, autofit);
        }

        /// <summary>
        /// Will insert bulk data in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="sheetName">Sheet to insert data. If the sheet is not found, it will be created</param>
        /// <param name="data">Data to insert in sheet</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="autofit">Should adjust all content columns width</param>
        public void InsertRows(string sheetName, IEnumerable<string[]> data, int rowIndex = 2, int columnIndex = 1, bool autofit = false)
        {
            var sheet = GetSheetByName(sheetName);
            InsertRows(sheet, data, rowIndex, columnIndex, autofit);
        }

        /// <summary>
        /// Will insert bulk data in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="sheet">Sheet to insert data</param>
        /// <param name="data">Data to insert in sheet</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="autofit">Should adjust all content columns width</param>
        public void InsertRows(ExcelWorksheet sheet, IEnumerable<string[]> data, int rowIndex = 2, int columnIndex = 1, bool autofit = false)
        {
            if (rowIndex <= 0) rowIndex = 1;
            if (columnIndex <= 0) columnIndex = 1;

            sheet.Cells[GetRangeForCell(rowIndex, columnIndex)].LoadFromArrays(data);

            if (autofit)
            {
                Autofit(sheet);
            }
        }

        /// <summary>
        /// Will create (if sheet is not already created) in the current document and set a header. Header auto-filter enabled for Excel
        /// </summary>
        /// <param name="sheetName">Sheet name to find or create</param>
        /// <param name="header">Header text values</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        /// <returns></returns>
        public ExcelWorksheet GetSheetByNameWithHeader(string sheetName, string[] header, int rowIndex = 1, int columnIndex = 1)
        {
            if (rowIndex <= 0) rowIndex = 1;
            if (columnIndex <= 0) columnIndex = 1;

            var sheet = GetSheetByName(sheetName);
            string range = $"{GetRangeForCell(rowIndex, columnIndex)}:{GetRangeForCell(rowIndex, header.Length)}";
            sheet.Cells[range].LoadFromArrays(new List<string[]>(new[] { header }));
            sheet.Cells[range].AutoFilter = true;
            sheet.Cells[range].Style.Font.Bold = true;
            sheet.Cells[range].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[range].Style.Fill.BackgroundColor.SetColor(Color.Black);
            sheet.Cells[range].AutoFitColumns();
            return sheet;
        }

        /// <summary>
        /// Will create (if sheet is not already created) in the current document
        /// </summary>
        /// <param name="sheetName">Sheet name to find or create</param>
        /// <returns></returns>
        public ExcelWorksheet GetSheetByName(string sheetName)
        {
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                var sheet = Workbook.Workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    sheet = Workbook.Workbook.Worksheets.Add(sheetName);
                }
                return sheet;
            }
            else
            {
                throw new ArgumentNullException("Argument sheetName is null or empty string");
            }
        }

        /// <summary>
        /// Enumerate sheets currently in the document
        /// </summary>
        /// <returns></returns>
        public string[] GetSheetNames()
        {
            return Workbook.Workbook.Worksheets.Select(x => x.Name).ToArray();
        }

        /// <summary>
        /// Sets width for all columns to adjust to their content
        /// </summary>
        /// <param name="sheet"></param>
        public void Autofit(ExcelWorksheet sheet)
        {
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }

        protected string GetRangeForCell(int rowIndex, int columnIndex)
        {
            return $"{GetExcelColumnName(columnIndex)}{rowIndex}";
        }

        protected string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulus;

            while(dividend > 0)
            {
                modulus = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulus).ToString();
                dividend = (int)((dividend - modulus) / 26);
            }

            return columnName;
        }

        protected string GetValidFileName(string fileName)
        {
            String ret = Regex.Replace(fileName.Trim(), "[^A-Za-z0-9_. ]+", "");
            return ret.Replace(' ', '_');
        }

        public void Dispose()
        {
            Workbook.Dispose();
        }
    }
}
