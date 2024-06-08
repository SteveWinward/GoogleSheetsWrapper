using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// Helper class for working with Google Sheets tabs
    /// </summary>
    public class SheetHelper
    {
        /// <summary>
        /// A Spreadsheet resource represents every spreadsheet and has a unique 
        /// spreadsheetId value, containing letters, numbers, hyphens, or underscores. 
        /// You can find the spreadsheet ID in a Google Sheets URL: 
        /// https://docs.google.com/spreadsheets/d/spreadsheetId/edit#gid=0
        /// </summary>
        public string SpreadsheetID { get; set; }

        /// <summary>
        /// The tab name of the Google Sheets document
        /// </summary>
        public string TabName { get; private set; }

        /// <summary>
        /// The Sheet ID for the Google Sheets document.  This value is set by the Google Sheets API.
        /// </summary>
        public int? SheetID { get; set; }

        /// <summary>
        /// The Google Service Account email address used for authentication
        /// </summary>
        public string ServiceAccountEmail { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string[] Scopes { get; set; } = { SheetsService.Scope.Spreadsheets };

        /// <summary>
        /// The Google API SheetsService object.  This is the core object we "wrap" CRUD operations against
        /// </summary>
        public SheetsService Service { get; private set; }

        /// <summary>
        /// Private field indicating if the Init() method has been executed
        /// </summary>
        private bool IsInitialized;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="spreadsheetID"></param>
        /// <param name="serviceAccountEmail"></param>
        /// <param name="tabName"></param>
        public SheetHelper(string spreadsheetID, string serviceAccountEmail, string tabName)
        {
            SpreadsheetID = spreadsheetID;
            ServiceAccountEmail = serviceAccountEmail;
            TabName = tabName;
        }

        /// <summary>
        /// Initializes the SheetHelper object
        /// </summary>
        /// <param name="jsonCredentials"></param>
        public void Init(string jsonCredentials)
        {
            Init(jsonCredentials, default);
        }

        /// <summary>
        /// Initializes the SheetHelper object with authentication
        /// </summary>
        /// <param name="jsonCredentials"></param>
        /// <param name="httpClientFactory"></param>
        public void Init(string jsonCredentials, Google.Apis.Http.IHttpClientFactory httpClientFactory)
        {
            var credential = (ServiceAccountCredential)
                   GoogleCredential.FromJson(jsonCredentials).UnderlyingCredential;

            // Authenticate as service account to the Sheets API
            var initializer = new ServiceAccountCredential.Initializer(credential.Id)
            {
                User = ServiceAccountEmail,
                Key = credential.Key,
                Scopes = Scopes,
                HttpClientFactory = httpClientFactory,
            };
            credential = new ServiceAccountCredential(initializer);

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                HttpClientFactory = httpClientFactory,
            });

            Service = service;
            IsInitialized = true;

            UpdateTabName(TabName);
        }

        /// <summary>
        /// Throws ArgumentException if the Init() method has not been called yet.
        /// </summary>
        /// <exception cref="ArgumentException"></exception>
        protected void EnsureServiceInitialized()
        {
            if (!IsInitialized)
            {
                throw new ArgumentException("SheetHelper requires the Init(string jsonCredentials) method to be called before using any of its methods.");
            }
        }

        /// <summary>
        /// Set the tab to the specified newTabName value
        /// </summary>
        /// <param name="newTabName"></param>
        public void UpdateTabName(string newTabName)
        {
            EnsureServiceInitialized();

            var spreadsheet = Service.Spreadsheets.Get(SpreadsheetID);

            var result = spreadsheet.Execute();

            Sheet sheet;

            // Lookup the sheet id for the given tab name
            if (!string.IsNullOrEmpty(newTabName))
            {
                if (!result.Sheets.Any(s => s.Properties.Title.Equals(newTabName, StringComparison.OrdinalIgnoreCase)))
                {
                    _ = CreateNewTab(newTabName);
                    result = spreadsheet.Execute();
                }

                sheet = result.Sheets.First(s => s.Properties.Title.Equals(newTabName, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                sheet = result.Sheets.First();
            }

            SheetID = sheet.Properties.SheetId;

            TabName = newTabName;
        }

        /// <summary>
        /// Returns a list of all tab names in the Google Spreadsheet
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllTabNames()
        {
            EnsureServiceInitialized();

            var spreadsheet = Service.Spreadsheets.Get(SpreadsheetID);

            var result = spreadsheet.Execute();

            var tabs = result.Sheets.Select(s => s.Properties.Title).ToList();

            return tabs;
        }

        /// <summary>
        /// Return a collection of rows for a given SheetRange input
        /// </summary>
        /// <param name="range"></param>
        /// <param name="valueRenderOption"></param>
        /// <param name="dateTimeRenderOption"></param>
        /// <returns></returns>
        public IList<IList<object>> GetRows(SheetRange range,
            ValueRenderOptionEnum valueRenderOption = ValueRenderOptionEnum.UNFORMATTEDVALUE,
            DateTimeRenderOptionEnum dateTimeRenderOption = DateTimeRenderOptionEnum.SERIALNUMBER)
        {
            EnsureServiceInitialized();

            var rangeValue = range.CanSupportA1Notation ? range.A1Notation : range.R1C1Notation;

            var request =
                    Service.Spreadsheets.Values.Get(SpreadsheetID, rangeValue);

            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption = dateTimeRenderOption;

            var response = request.Execute();

            if (response.Values != null)
            {
                return response.Values;
            }
            else
            {
                return new List<IList<object>>();
            }
        }

        /// <summary>
        /// Return a collection of rows formatted values for a given SheetRange input
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public IList<IList<object>> GetRowsFormatted(SheetRange range)
        {
            EnsureServiceInitialized();

            var rangeValue = range.CanSupportA1Notation ? range.A1Notation : range.R1C1Notation;

            var request =
                    Service.Spreadsheets.Values.Get(SpreadsheetID, rangeValue);

            request.ValueRenderOption = ValueRenderOptionEnum.FORMATTEDVALUE;
            request.DateTimeRenderOption = DateTimeRenderOptionEnum.FORMATTEDSTRING;

            var response = request.Execute();

            if (response.Values != null)
            {
                return response.Values;
            }
            else
            {
                return new List<IList<object>>();
            }
        }

        /// <summary>
        /// Clears values from a spreadsheet (NOTE: All other properties of the cell (such as formatting, data validation, etc..) are kept.)
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public ClearValuesResponse ClearRange(SheetRange range)
        {
            EnsureServiceInitialized();

            var rangeValue = range.CanSupportA1Notation ? range.A1Notation : range.R1C1Notation;

            var requestBody = new ClearValuesRequest();

            var clearRequest = Service.Spreadsheets.Values.Clear(requestBody, SpreadsheetID, rangeValue);
            return clearRequest.Execute();
        }

        /// <summary>
        /// Deletes a specified column
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse DeleteColumn(int col)
        {
            EnsureServiceInitialized();

            var request = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        Dimension = "COLUMNS",
                        StartIndex = col - 1,
                        EndIndex = col,
                        SheetId = SheetID,
                    }
                }
            };

            var bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = Service.Spreadsheets.BatchUpdate(bussr, SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Deletes a specified column
        /// </summary>
        /// <param name="columnLetter"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse DeleteColumn(string columnLetter)
        {
            EnsureServiceInitialized();

            var columnId = SheetRange.GetColumnIDFromLetters(columnLetter);

            return DeleteColumn(columnId);
        }

        /// <summary>
        /// Deletes a specified row
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse DeleteRow(int row)
        {
            EnsureServiceInitialized();

            var request = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        Dimension = "ROWS",
                        StartIndex = row - 1,
                        EndIndex = row,
                        SheetId = SheetID,
                    }
                }
            };

            var bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = Service.Spreadsheets.BatchUpdate(bussr, SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse DeleteRows(int startRow, int endRow)
        {
            EnsureServiceInitialized();

            var request = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        Dimension = "ROWS",
                        StartIndex = startRow - 1,
                        EndIndex = endRow,
                        SheetId = SheetID,
                    }
                }
            };

            var bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = Service.Spreadsheets.BatchUpdate(bussr, SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Inserts a blank new column using the column index as the id (NOTE: 1 is the first index for the column based index)
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse InsertBlankColumn(int column)
        {
            EnsureServiceInitialized();

            if (column < 1)
            {
                throw new ArgumentException("column index value must be 1 or greater");
            }

            var request = new Request()
            {
                InsertDimension = new InsertDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        Dimension = "COLUMNS",
                        StartIndex = column - 1,
                        EndIndex = column,
                        SheetId = SheetID,
                    },
                    InheritFromBefore = column > 1,
                }
            };

            var bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = Service.Spreadsheets.BatchUpdate(bussr, SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Inserts a blank new column using a letter notation (i.e. B2 as the column id)
        /// </summary>
        /// <param name="columnLetter"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse InsertBlankColumn(string columnLetter)
        {
            EnsureServiceInitialized();

            var columnId = SheetRange.GetColumnIDFromLetters(columnLetter);

            return InsertBlankColumn(columnId);
        }

        /// <summary>
        /// Inserts a new blank row
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse InsertBlankRow(int row)
        {
            EnsureServiceInitialized();

            if (row < 1)
            {
                throw new ArgumentException("row index value must be 1 or greater");
            }

            var request = new Request()
            {
                InsertDimension = new InsertDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        Dimension = "ROWS",
                        StartIndex = row - 1,
                        EndIndex = row,
                        SheetId = SheetID,
                    },
                    InheritFromBefore = row > 1,
                }
            };

            var bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = Service.Spreadsheets.BatchUpdate(bussr, SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Runs a collection of updates as a batch operation in a single call.
        ///
        /// This is useful to avoid throttling limits with the Google Sheets API
        /// </summary>
        /// <param name="updates"></param>
        /// <param name="fieldMask">Allows you to specify what fields you want to update in the BatchUpdate call,
        /// defaults to userEnteredValue to keep existing cell styles, use "*" to update all properties here.
        /// Other valid field mask values are: dataSourceFormula, dataSourceTable, dataValidation, effectiveFormat, effectiveValue, formattedValue, hyperlink, note, pivotTable, textFormatRuns, userEnteredFormat, userEnteredValue
        /// </param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse BatchUpdate(List<BatchUpdateRequestObject> updates, string fieldMask = "userEnteredValue")
        {
            EnsureServiceInitialized();

            var bussr = new BatchUpdateSpreadsheetRequest();

            var requests = new List<Request>(updates.Count);

            foreach (var update in updates)
            {
                var gridRange = new GridRange()
                {
                    SheetId = SheetID,
                    StartColumnIndex = update.Range.StartColumn - 1,
                    StartRowIndex = update.Range.StartRow - 1,
                };

                // If this is a single cell we use the start column and start row to determine the end range
                if (update.Range.IsSingleCellRange)
                {
                    gridRange.EndColumnIndex = update.Range.StartColumn;
                    gridRange.EndRowIndex = update.Range.StartRow;
                }
                // if this is a range, we can use the actual end column and end row information for the end range
                else
                {
                    gridRange.EndColumnIndex = update.Range.EndColumn;
                    gridRange.EndRowIndex = update.Range.EndRow;
                }

                //create the update request for cells from the first row
                var updateCellsRequest = new Request()
                {
                    RepeatCell = new RepeatCellRequest()
                    {
                        Range = gridRange,
                        Cell = update.Data,
                        Fields = fieldMask
                    }
                };

                requests.Add(updateCellsRequest);
            }

            bussr.Requests = requests;

            var updateRequest = Service.Spreadsheets.BatchUpdate(bussr, SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Create a blank new tab to the specified newTabName value
        /// </summary>
        /// <param name="newTabName"></param>
        private BatchUpdateSpreadsheetResponse CreateNewTab(string newTabName)
        {
            EnsureServiceInitialized();

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>
                {
                    new Request
                    {
                        AddSheet = new AddSheetRequest
                        {
                            Properties = new SheetProperties
                            {
                                Title = newTabName,
                                Index = 0
                            }
                        }
                    }
                }
            };

            var batchUpdateRequest = Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SpreadsheetID);

            return batchUpdateRequest.Execute();
        }
    }

    /// <summary>
    /// SheetHelper strongly typed class
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class SheetHelper<T> : SheetHelper where T : BaseRecord
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="spreadsheetID"></param>
        /// <param name="serviceAccountEmail"></param>
        /// <param name="tabName"></param>
        public SheetHelper(string spreadsheetID, string serviceAccountEmail, string tabName)
            : base(spreadsheetID, serviceAccountEmail, tabName) { }

        /// <summary>
        /// Adds a record to the next row in the Google Sheet tab
        /// </summary>
        /// <param name="record"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse AppendRow(T record)
        {
            EnsureServiceInitialized();

            return AppendRows(new T[] { record });
        }

        /// <summary>
        /// Adds mulitlpe rows to the next row in the Google Sheets tab
        /// </summary>
        /// <param name="records"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse AppendRows(IList<T> records)
        {
            EnsureServiceInitialized();

            var rows = new List<RowData>(records.Count);

            foreach (var record in records)
            {
                if (record != null)
                {
                    var row = new RowData
                    {
                        Values = record.ConvertToCellData(TabName).Select(b => b.Data).ToList(),
                    };

                    rows.Add(row);
                }
            }

            var appendRequest = new AppendCellsRequest
            {
                Fields = "*",
                SheetId = SheetID,
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
            return Service.Spreadsheets
                .BatchUpdate(batchRequst, SpreadsheetID)
                .Execute();
        }
    }
}