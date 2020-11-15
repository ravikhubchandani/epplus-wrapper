using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace SpreadsheetWrapper
{
    /// <summary>
    /// Build and export spreadsheets to Excel, CSV or Byte Array
    /// </summary>
    public class SpreadsheetWrapper : IDisposable
    {
        public ExcelPackage Workbook { get; private set; }

        public SpreadsheetWrapper()
        {
            Workbook = new ExcelPackage();
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
        /// Save the specified sheet as CSV text file
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="fileName"></param>
        /// <param name="basePath"></param>
        /// <param name="separator"></param>
        public void SaveCsvAs(string sheetName, string fileName, string basePath = @".\", char separator = ';')
        {
            var sheet = GetSheetByName(sheetName);
            SaveCsvAs(sheet, fileName, basePath, separator);
        }

        /// <summary>
        /// Save the specified sheet as CSV UTF-8 encoded text file
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fileName"></param>
        /// <param name="basePath"></param>
        /// <param name="separator"></param>
        public void SaveCsvAs(ExcelWorksheet sheet, string fileName, string basePath = @".\", char separator = ';')
        {
            string filePath = Path.Combine(basePath, fileName);
            string content = GetSheetAsCsvString(sheet, separator);
            File.WriteAllText(filePath, content, Encoding.UTF8);
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
        /// <param name="sheetName">Sheet to insert data. If the sheet is not found, it will be created</param>
        /// <param name="data">Data to insert in sheet</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        public void InsertRows(string sheetName, IEnumerable<object[]> data, int rowIndex = 2, int columnIndex = 1)
        {
            var sheet = GetSheetByName(sheetName);
            InsertRows(sheet, data, rowIndex, columnIndex);
        }

        /// <summary>
        /// Will insert bulk data in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="sheet">Sheet to insert data</param>
        /// <param name="data">Data to insert in sheet</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        public void InsertRows(ExcelWorksheet sheet, IEnumerable<object[]> data, int rowIndex = 2, int columnIndex = 1)
        {
            sheet.Cells[GetRangeForCell(rowIndex, columnIndex)].LoadFromArrays(data);
        }

        /// <summary>
        /// Will insert data parameter in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="sheetName">Sheet to insert data. If the sheet is not found, it will be created</param>
        /// <param name="data">Data to insert in sheet</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        public void InsertRow(string sheetName, object[] data, int rowIndex = 2, int columnIndex = 1)
        {
            var sheet = GetSheetByName(sheetName);
            InsertRow(sheet, data, rowIndex, columnIndex);
        }

        /// <summary>
        /// Will insert data parameter in the specified sheet starting at the specified index
        /// </summary>
        /// <param name="sheet">Sheet to insert data</param>
        /// <param name="data">Data to insert in sheet</param>
        /// <param name="rowIndex">Row index to start inserting. Index starts from 1</param>
        /// <param name="columnIndex">Row index to start inserting. Index starts from 1</param>
        public void InsertRow(ExcelWorksheet sheet, object[] data, int rowIndex = 2, int columnIndex = 1)
        {
            sheet.Cells[GetRangeForCell(rowIndex, columnIndex)].LoadFromCollection(data);
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
            var sheet = GetSheetByName(sheetName);
            string endColumn = GetExcelColumnName(header.Length);
            string range = GetRangeForCell(rowIndex, columnIndex);
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
            var sheet = Workbook.Workbook.Worksheets[sheetName];
            if(sheet == null)
            {
                sheet = Workbook.Workbook.Worksheets.Add(sheetName);
            }
            return sheet;
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

        public void Dispose()
        {
            Workbook.Dispose();
        }
    }
}
