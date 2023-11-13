using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using CsvHelper;
using CsvHelper.Configuration;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// Helper class for appender data to Google Sheets tab
    /// </summary>
    public class SheetAppender
    {
        private readonly SheetHelper _sheetHelper;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="spreadsheetID"></param>
        /// <param name="serviceAccountEmail"></param>
        /// <param name="tabName"></param>
        public SheetAppender(string spreadsheetID, string serviceAccountEmail, string tabName)
        {
            this._sheetHelper = new SheetHelper(spreadsheetID, serviceAccountEmail, tabName);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sheetHelper"></param>
        public SheetAppender(SheetHelper sheetHelper)
        {
            this._sheetHelper = sheetHelper;
        }

        /// <summary>
        /// Initializes the SheetAppender object with authentication to Google Sheets API
        /// </summary>
        /// <param name="jsonCredentials"></param>
        public void Init(string jsonCredentials)
        {
            this._sheetHelper.Init(jsonCredentials);
        }

        /// <summary>
        /// Appends a CSV file and all its rows into the current Google Sheets tab
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="includeHeaders"></param>
        /// <param name="batchWaitTime"></param>
        public void AppendCsv(string filePath, bool includeHeaders, int batchWaitTime = 1000)
        {
#pragma warning disable IDE0063 // Using statement can be simplified

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                this.AppendCsv(stream, includeHeaders, batchWaitTime);
            }

#pragma warning restore IDE0063 // Using statement can be simplified
        }

        /// <summary>
        /// Appends a CSV file and all its rows into the current Google Sheets tab
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="includeHeaders"></param>
        /// <param name="batchWaitTime"></param>
        public void AppendCsv(Stream stream, bool includeHeaders, int batchWaitTime = 1000)
        {
            using var streamReader = new StreamReader(stream);
            using var csv = new CsvReader(streamReader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = includeHeaders
            });
            {
                var batchRowLimit = 100;

                var dataRecords =
                    csv.GetRecords<dynamic>()
                    .Select(x => (IDictionary<string, object>)x)
                    .ToList();

                var rowData = new List<RowData>();

                var currentBatchCount = 0;

                if (includeHeaders)
                {
                    currentBatchCount++;

                    var row = new RowData()
                    {
                        Values = new List<CellData>()
                    };

                    var headers = csv.HeaderRecord;

                    foreach (var header in headers)
                    {
                        row.Values.Add(this.StringToCellData(header));
                    }

                    rowData.Add(row);
                }

                foreach (var record in dataRecords)
                {
                    currentBatchCount++;

                    var row = new RowData()
                    {
                        Values = new List<CellData>()
                    };

                    foreach (var property in record.Keys)
                    {
                        row.Values.Add(this.StringToCellData(record[property]));
                    }

                    rowData.Add(row);

                    if (currentBatchCount >= batchRowLimit)
                    {
                        this.AppendRows(rowData);

                        rowData = new List<RowData>();
                        currentBatchCount = 0;

                        Thread.Sleep(batchWaitTime);
                    }
                }

                if (rowData.Count > 0)
                {
                    this.AppendRows(rowData);
                }
            }
        }

        /// <summary>
        /// Append a single row to the spreadsheet
        /// </summary>
        /// <param name="row"></param>
        public void AppendRow(RowData row)
        {
            this.AppendRows(new List<RowData>() { row });
        }

        /// <summary>
        /// Append a single weakly typed row of string data to the spreadsheet
        /// </summary>
        /// <param name="rowStringValues"></param>
        public void AppendRow(List<string> rowStringValues)
        {
            this.AppendRows(new List<List<string>>() { rowStringValues });
        }

        /// <summary>
        /// Append multiple rows of data to the spreadsheet
        /// </summary>
        /// <param name="rows"></param>
        public void AppendRows(List<RowData> rows)
        {
            var appendRequest = new AppendCellsRequest
            {
                Fields = "*",
                SheetId = this._sheetHelper.SheetID,
                Rows = rows
            };

            var request = new Request
            {
                AppendCells = appendRequest
            };

            // Wrap it into batch update request.
            var batchRequst = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            // Finally update the sheet.
            _ = this._sheetHelper.Service.Spreadsheets
                .BatchUpdate(batchRequst, this._sheetHelper.SpreadsheetID)
                .Execute();
        }

        /// <summary>
        /// Append multilpe weakly typed string data rows to the spreadsheet
        /// </summary>
        /// <param name="stringRows"></param>
        public void AppendRows(List<List<string>> stringRows)
        {
            var rows = new List<RowData>();

            foreach (var stringRow in stringRows)
            {
                var row = new RowData()
                {
                    Values = new List<CellData>(),
                };

                foreach (var stringValue in stringRow)
                {
                    row.Values.Add(this.StringToCellData(stringValue));
                }

                rows.Add(row);
            }

            this.AppendRows(rows);
        }

        /// <summary>
        /// Converts a string into its appropriate cell data object
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private CellData StringToCellData(object value)
        {
            var cell = new CellData
            {
                UserEnteredValue = new ExtendedValue()
            };

            if (value != null)
            {
                cell.UserEnteredValue.StringValue = value.ToString();
            }
            else
            {
                cell.UserEnteredValue.StringValue = string.Empty;
            }

            return cell;
        }
    }
}