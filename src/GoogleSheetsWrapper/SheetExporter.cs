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
        /// <param name="spreadsheetID">The identifier of the spreadsheet to export.</param>
        /// <param name="serviceAccountEmail">The service account email used for authentication.</param>
        /// <param name="tabName">The tab to export.</param>
        public SheetExporter(string spreadsheetID, string serviceAccountEmail, string tabName)
        {
            _sheetHelper = new SheetHelper(spreadsheetID, serviceAccountEmail, tabName);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sheetHelper">The initialized sheet helper to use for export operations.</param>
        public SheetExporter(SheetHelper sheetHelper)
        {
            _sheetHelper = sheetHelper;
        }

        /// <summary>
        /// Initializes the underlying sheet helper using service account credentials.
        /// </summary>
        /// <param name="jsonCredentials">Service account credentials in JSON format.</param>
        public void Init(string jsonCredentials)
        {
            _sheetHelper.Init(jsonCredentials);
        }

        /// <summary>
        /// Exports the current Google Sheet tab to a CSV file
        /// </summary>
        /// <param name="range">The range to export.</param>
        /// <param name="stream">The destination stream for the CSV content.</param>
        /// <param name="delimiter">The delimiter used between CSV fields.</param>
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
        /// <param name="range">The range to export.</param>
        /// <param name="stream">The destination stream for the CSV content.</param>
        /// <param name="csvConfiguration">The CsvHelper settings used to generate the CSV content.</param>
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
        /// Exports the specified Google Sheet range to an Excel workbook.
        /// </summary>
        /// <param name="range">The range to export.</param>
        /// <param name="stream">The destination stream for the Excel workbook.</param>
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