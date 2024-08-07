using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// Helper class to export data from Google Sheets
    /// </summary>
    public class SheetExporter
    {
        private readonly SheetHelper _sheetHelper;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="spreadsheetID"></param>
        /// <param name="serviceAccountEmail"></param>
        /// <param name="tabName"></param>
        public SheetExporter(string spreadsheetID, string serviceAccountEmail, string tabName)
        {
            _sheetHelper = new SheetHelper(spreadsheetID, serviceAccountEmail, tabName);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sheetHelper"></param>
        public SheetExporter(SheetHelper sheetHelper)
        {
            _sheetHelper = sheetHelper;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="jsonCredentials"></param>
        public void Init(string jsonCredentials)
        {
            _sheetHelper.Init(jsonCredentials);
        }

        /// <summary>
        /// Exports the current Google Sheet tab to a CSV file
        /// </summary>
        /// <param name="range"></param>
        /// <param name="stream"></param>
        /// <param name="delimiter"></param>
        public void ExportAsCsv(SheetRange range, Stream stream, string delimiter = ",")
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter,
            };

            ExportAsCsv(range, stream, config);
        }

        /// <summary>
        /// Exports the current Google Sheet tab to a CSV file.  This override lets you explicitly specify the CsvConfiguration object for the CsvHelper library.
        /// </summary>
        /// <param name="range"></param>
        /// <param name="stream"></param>
        /// <param name="csvConfiguration"></param>
        public void ExportAsCsv(SheetRange range, Stream stream, CsvConfiguration csvConfiguration)
        {
            var rows = _sheetHelper.GetRowsFormatted(range);

            using var streamWriter = new StreamWriter(stream);
            using var csv = new CsvWriter(streamWriter, csvConfiguration);
            foreach (var row in rows)
            {
                foreach (var cell in row)
                {
                    csv.WriteField(cell?.ToString());
                }

                csv.NextRecord();
            }
        }

        /// <summary>
        /// Exports the current Google Sheet tab to a CSV file
        /// </summary>
        /// <param name="range"></param>
        /// <param name="stream"></param>
        public void ExportAsExcel(SheetRange range, Stream stream)
        {
            var rows = _sheetHelper.GetRowsFormatted(range);

            var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            var workbookpart = document.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            // Add Sheets to the Workbook.
            var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            var sheet = new Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = string.IsNullOrEmpty(_sheetHelper.TabName) ? "Sheet1" : _sheetHelper.TabName
            };
            sheets.Append(sheet);

            foreach (var row in rows)
            {
                var newRow = new Row();

                foreach (var value in row)
                {
                    var cell = CreateCell(value?.ToString());
                    _ = newRow.AppendChild(cell);
                }

                _ = sheetData.AppendChild(newRow);
            }

            workbookpart.Workbook.Save();

            document.Dispose();
        }

        private static Cell CreateCell(string text)
        {
            var cell = new Cell
            {
                DataType = ResolveCellDataTypeOnValue(text),
                CellValue = new CellValue(text)
            };

            return cell;
        }

        private static EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            if (int.TryParse(text, out _) || double.TryParse(text, out _))
            {
                return (EnumValue<CellValues>)CellValues.Number;
            }
            else
            {
                return (EnumValue<CellValues>)CellValues.String;
            }
        }
    }
}