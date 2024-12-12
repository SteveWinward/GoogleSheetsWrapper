using System;
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
            _sheetHelper = new SheetHelper(spreadsheetID, serviceAccountEmail, tabName);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sheetHelper"></param>
        public SheetAppender(SheetHelper sheetHelper)
        {
            _sheetHelper = sheetHelper;
        }

        /// <summary>
        /// Initializes the SheetAppender object with authentication to Google Sheets API
        /// </summary>
        /// <param name="jsonCredentials"></param>
        public void Init(string jsonCredentials)
        {
            _sheetHelper.Init(jsonCredentials);
        }

        /// <summary>
        /// Appends a CSV file and all its rows into the current Google Sheets tab
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="csvHasHeaderRecord">This boolean indicates whether the CSV file has a header record row or not</param>
        /// <param name="batchWaitTime">See https://developers.google.com/sheets/api/limits at last check is 60 requests a minute, so 1 second delay per request should avoid limiting</param>
        /// <param name="batchSize">Increasing batch size may improve throughput. Default is conservative.</param>
        /// <param name="skipWritingHeaderRow">This boolean indicates if you want to actually write the header row to the Google sheet</param>
        public void AppendCsv(string filePath, bool csvHasHeaderRecord, int batchWaitTime = 1000, int batchSize = 100, bool skipWritingHeaderRow = false)
        {
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                AppendCsv(stream, csvHasHeaderRecord, batchWaitTime, batchSize, skipWritingHeaderRow);
            }
        }

        /// <summary>
        /// Appends a CSV file and all its rows into the current Google Sheets tab
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="csvHasHeaderRecord">This boolean indicates whether the CSV file has a header record row or not</param>
        /// <param name="batchWaitTime">See https://developers.google.com/sheets/api/limits at last check is 60 requests a minute, so 1 second delay per request should avoid limiting</param>
        /// <param name="batchSize">Increasing batch size may improve throughput. Default is conservative.</param>
        /// <param name="skipWritingHeaderRow">This boolean indicates if you want to actually write the header row to the Google sheet</param>
        public void AppendCsv(Stream stream, bool csvHasHeaderRecord, int batchWaitTime = 1000, int batchSize = 100, bool skipWritingHeaderRow = false)
        {
            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = csvHasHeaderRecord
            };

            AppendCsv(stream, csvConfig, batchWaitTime, batchSize, skipWritingHeaderRow);
        }

        /// <summary>
        /// Appends a CSV file and all its rows into the current Google Sheets tab
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="csvConfig"></param>
        /// <param name="batchWaitTime">See https://developers.google.com/sheets/api/limits at last check is 60 requests a minute, so 1 second delay per request should avoid limiting</param>
        /// <param name="batchSize">Increasing batch size may improve throughput. Default is conservative.</param>
        /// <param name="skipWritingHeaderRow">This boolean indicates if you want to actually write the header row to the Google sheet</param>
        public void AppendCsv(Stream stream, CsvConfiguration csvConfig, int batchWaitTime = 1000, int batchSize = 100, bool skipWritingHeaderRow = false)
        {
            using var streamReader = new StreamReader(stream);
            using var csv = new CsvReader(streamReader, csvConfig);
            {
                var batchRowLimit = batchSize;

                var dataRecords =
                    csv.GetRecords<dynamic>()
                    .Select(x => (IDictionary<string, object>)x)
                    .ToList();

                var rowData = new List<RowData>();

                var currentBatchCount = 0;

                // Only write the header record if its specified to not skip writing the header row
                if (csvConfig.HasHeaderRecord && !skipWritingHeaderRow)
                {
                    currentBatchCount++;

                    var row = new RowData()
                    {
                        Values = new List<CellData>()
                    };

                    var headers = csv.HeaderRecord;

                    foreach (var header in headers)
                    {
                        row.Values.Add(StringToCellData(header));
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
                        row.Values.Add(StringToCellData(record[property]));
                    }

                    rowData.Add(row);

                    if (currentBatchCount >= batchRowLimit)
                    {
                        AppendRows(rowData);

                        rowData = new List<RowData>();
                        currentBatchCount = 0;

                        Thread.Sleep(batchWaitTime);
                    }
                }

                if (rowData.Count > 0)
                {
                    AppendRows(rowData);
                }
            }
        }

        /// <summary>
        /// This lets you append a list of objects to the Google Sheet with any object using reflection to determine the property types during runtime.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataRecords">The list of records</param>
        /// <param name="batchWaitTime">See https://developers.google.com/sheets/api/limits at last check is 60 requests a minute, so 1 second delay per request should avoid limiting</param>
        /// <param name="batchSize">Increasing batch size may improve throughput. Default is conservative.</param>
        /// <param name="skipWritingHeaderRow">This boolean indicates if you want to actually write the header row to the Google sheet</param>
        public void AppendObject<T>(IEnumerable<T> dataRecords, int batchWaitTime = 1000, int batchSize = 100, bool skipWritingHeaderRow = false) where T : class
        {
            var batchRowLimit = batchSize;

            var rowData = new List<RowData>();

            var currentBatchCount = 0;

            var properties = typeof(T).GetProperties().Where(x => !x.CustomAttributes.Any(a => a.AttributeType == typeof(CsvHelper.Configuration.Attributes.IgnoreAttribute)));

            // Only write the header record if its specified to not skip writing the header row
            if (!skipWritingHeaderRow)
            {
                currentBatchCount++;

                var row = new RowData()
                {
                    Values = new List<CellData>()
                };

                foreach (var header in properties)
                {
                    var name = header.Name;
                    var nameAttribute = header.GetCustomAttributes(typeof(CsvHelper.Configuration.Attributes.NameAttribute), true).FirstOrDefault();
                    if (nameAttribute != null)
                    {
                        name = (nameAttribute as CsvHelper.Configuration.Attributes.NameAttribute).Names.FirstOrDefault() ?? name;
                    }
                    row.Values.Add(StringToCellData(name));
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

                foreach (var property in properties)
                {
                    var propertyValue = property.GetValue(record);
                    row.Values.Add(ObjectToCellData(propertyValue));
                }

                rowData.Add(row);

                if (currentBatchCount >= batchRowLimit)
                {
                    AppendRows(rowData);

                    rowData = new List<RowData>();
                    currentBatchCount = 0;

                    Thread.Sleep(batchWaitTime);
                }
            }

            if (rowData.Count > 0)
            {
                AppendRows(rowData);
            }
        }

        /// <summary>
        /// Append a single row to the spreadsheet
        /// </summary>
        /// <param name="row"></param>
        public void AppendRow(RowData row)
        {
            AppendRows(new List<RowData>() { row });
        }

        /// <summary>
        /// Append a single weakly typed row of string data to the spreadsheet
        /// </summary>
        /// <param name="rowStringValues"></param>
        public void AppendRow(List<string> rowStringValues)
        {
            AppendRows(new List<List<string>>() { rowStringValues });
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
                SheetId = _sheetHelper.SheetID,
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
            _ = _sheetHelper.Service.Spreadsheets
                .BatchUpdate(batchRequst, _sheetHelper.SpreadsheetID)
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
                    row.Values.Add(StringToCellData(stringValue));
                }

                rows.Add(row);
            }

            AppendRows(rows);
        }

        /// <summary>
        /// Converts a string into its appropriate cell data object
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private static CellData StringToCellData(object value)
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

        /// <summary>
        /// Converts a object into its appropriate cell data object
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private static CellData ObjectToCellData(object value)
        {
            var cell = new CellData
            {
                UserEnteredValue = new ExtendedValue()
            };

            if (value != null)
            {
                if (value is bool b)
                {
                    cell.UserEnteredValue.BoolValue = b;
                }
                else if (value is int i)
                {
                    cell.UserEnteredValue.NumberValue = i;
                }
                else if (value is double dbl)
                {
                    cell.UserEnteredValue.NumberValue = dbl;
                }
                else if (value is decimal dec)
                {
                    cell.UserEnteredValue.NumberValue = decimal.ToDouble(dec);
                }
                else if (value is DateTime dt)
                {
                    cell.UserEnteredFormat = new CellFormat { NumberFormat = new NumberFormat { Type = "Date" } };
                    cell.UserEnteredValue.NumberValue = dt.ToOADate();
                }
                else
                {
                    cell.UserEnteredValue.StringValue = value.ToString();
                }
            }
            else
            {
                cell.UserEnteredValue.StringValue = string.Empty;
            }

            return cell;
        }
    }
}