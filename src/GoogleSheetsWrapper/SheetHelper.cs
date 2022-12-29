using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;

namespace GoogleSheetsWrapper
{
    public class SheetHelper
    {
        public string SpreadsheetID { get; set; }

        public string TabName { get; private set; }

        public int? SheetID { get; set; }

        public string ServiceAccountEmail { get; set; }

        public string[] Scopes { get; set; } = { SheetsService.Scope.Spreadsheets };

        public SheetsService Service { get; private set; }

        public SheetHelper(string spreadsheetID, string serviceAccountEmail, string tabName)
        {
            this.SpreadsheetID = spreadsheetID;
            this.ServiceAccountEmail = serviceAccountEmail;
            this.TabName = tabName;
        }
        public void Init(string jsonCredentials) => Init(jsonCredentials, default);
        public void Init(string jsonCredentials, Google.Apis.Http.IHttpClientFactory httpClientFactory)
        {
            var credential = (ServiceAccountCredential)
                   GoogleCredential.FromJson(jsonCredentials).UnderlyingCredential;

            // Authenticate as service account to the Sheets API
            var initializer = new ServiceAccountCredential.Initializer(credential.Id)
            {
                User = this.ServiceAccountEmail,
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

            this.Service = service;

            this.UpdateTabName(this.TabName);
        }

        /// <summary>
        /// Set the tab to the specified newTabName value
        /// </summary>
        /// <param name="newTabName"></param>
        public void UpdateTabName(string newTabName)
        {
            var spreadsheet = this.Service.Spreadsheets.Get(this.SpreadsheetID);

            var result = spreadsheet.Execute();

            Sheet sheet;

            // Lookup the sheet id for the given tab name
            if (!string.IsNullOrEmpty(newTabName))
            {
                if (!result.Sheets.Any(s => s.Properties.Title.Equals(newTabName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    CreateNewTab(newTabName);
                    result = spreadsheet.Execute();
                }

                sheet = result.Sheets.First(s => s.Properties.Title.Equals(newTabName, StringComparison.CurrentCultureIgnoreCase));
            }
            else
            {
                sheet = result.Sheets.First();
            }

            this.SheetID = sheet.Properties.SheetId;

            this.TabName = newTabName;
        }

        /// <summary>
        /// Returns a list of all tab names in the Google Spreadsheet
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllTabNames()
        {
            var spreadsheet = this.Service.Spreadsheets.Get(this.SpreadsheetID);

            var result = spreadsheet.Execute();

            var tabs = result.Sheets.Select(s => s.Properties.Title).ToList();

            return tabs;
        }

        /// <summary>
        /// Return a collection of rows for a given SheetRange input
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public IList<IList<object>> GetRows(SheetRange range,
            ValueRenderOptionEnum valueRenderOption = ValueRenderOptionEnum.UNFORMATTEDVALUE,
            DateTimeRenderOptionEnum dateTimeRenderOption = DateTimeRenderOptionEnum.SERIALNUMBER)
        {
            var rangeValue = range.CanSupportA1Notation ? range.A1Notation : range.R1C1Notation;

            GetRequest request =
                    this.Service.Spreadsheets.Values.Get(this.SpreadsheetID, rangeValue);

            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption = dateTimeRenderOption;

            ValueRange response = request.Execute();
            return response.Values;
        }

        /// <summary>
        /// Return a collection of rows formatted values for a given SheetRange input
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public IList<IList<object>> GetRowsFormatted(SheetRange range)
        {
            var rangeValue = range.CanSupportA1Notation ? range.A1Notation : range.R1C1Notation;

            GetRequest request =
                    this.Service.Spreadsheets.Values.Get(this.SpreadsheetID, rangeValue);

            request.ValueRenderOption = GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;
            request.DateTimeRenderOption = GetRequest.DateTimeRenderOptionEnum.FORMATTEDSTRING;

            ValueRange response = request.Execute();
            return response.Values;
        }

        /// <summary>
        /// Clears values from a spreadsheet (NOTE: All other properties of the cell (such as formatting, data validation, etc..) are kept.)
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public ClearValuesResponse ClearRange(SheetRange range)
        {
            var rangeValue = range.CanSupportA1Notation ? range.A1Notation : range.R1C1Notation;

            var requestBody = new ClearValuesRequest();

            var clearRequest = this.Service.Spreadsheets.Values.Clear(requestBody, this.SpreadsheetID, rangeValue);
            return clearRequest.Execute();
        }

        /// <summary>
        /// Deletes a specified column
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse DeleteColumn(int col)
        {
            var request = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        Dimension = "COLUMNS",
                        StartIndex = col - 1,
                        EndIndex = col,
                        SheetId = this.SheetID,
                    }
                }
            };

            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = this.Service.Spreadsheets.BatchUpdate(bussr, this.SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Deletes a specified column
        /// </summary>
        /// <param name="columnLetter"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse DeleteColumn(string columnLetter)
        {
            var columnId = SheetRange.GetColumnIDFromLetters(columnLetter);

            return this.DeleteColumn(columnId);
        }

        /// <summary>
        /// Deletes a specified row
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse DeleteRow(int row)
        {
            var request = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        Dimension = "ROWS",
                        StartIndex = row - 1,
                        EndIndex = row,
                        SheetId = this.SheetID,
                    }
                }
            };

            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = this.Service.Spreadsheets.BatchUpdate(bussr, this.SpreadsheetID);
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
            var request = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        Dimension = "ROWS",
                        StartIndex = startRow - 1,
                        EndIndex = endRow,
                        SheetId = this.SheetID,
                    }
                }
            };

            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = this.Service.Spreadsheets.BatchUpdate(bussr, this.SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Inserts a blank new column using the column index as the id (NOTE: 1 is the first index for the column based index)
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse InsertBlankColumn(int column)
        {
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
                        SheetId = this.SheetID,
                    },
                    InheritFromBefore = column > 0,
                }
            };

            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = this.Service.Spreadsheets.BatchUpdate(bussr, this.SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Inserts a blank new column using a letter notation (i.e. B2 as the column id)
        /// </summary>
        /// <param name="columnLetter"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse InsertBlankColumn(string columnLetter)
        {
            var columnId = SheetRange.GetColumnIDFromLetters(columnLetter);

            return this.InsertBlankColumn(columnId);
        }

        /// <summary>
        /// Inserts a new blank row
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse InsertBlankRow(int row)
        {
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
                        SheetId = this.SheetID,
                    },
                    InheritFromBefore = row > 0,
                }
            };

            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            var updateRequest = this.Service.Spreadsheets.BatchUpdate(bussr, this.SpreadsheetID);
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
            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest();

            var requests = new List<Request>(updates.Count);

            foreach (var update in updates)
            {
                var gridRange = new GridRange()
                {
                    SheetId = this.SheetID,
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

            var updateRequest = this.Service.Spreadsheets.BatchUpdate(bussr, this.SpreadsheetID);
            return updateRequest.Execute();
        }

        /// <summary>
        /// Create a blank new tab to the specified newTabName value
        /// </summary>
        /// <param name="newTabName"></param>
        private BatchUpdateSpreadsheetResponse CreateNewTab(string newTabName)
        {
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

            var batchUpdateRequest = this.Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, this.SpreadsheetID);

            return batchUpdateRequest.Execute();
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class SheetHelper<T> : SheetHelper where T : BaseRecord
    {
        public SheetHelper(string spreadsheetID, string serviceAccountEmail, string tabName)
            : base(spreadsheetID, serviceAccountEmail, tabName) { }

        public BatchUpdateSpreadsheetResponse AppendRow(T record)
        {
            return this.AppendRows(new T[] { record });
        }

        public BatchUpdateSpreadsheetResponse AppendRows(IList<T> records)
        {
            var rows = new List<RowData>(records.Count);

            foreach (var record in records)
            {
                var row = new RowData
                {
                    Values = record.ConvertToCellData(this.TabName).Select(b => b.Data).ToList(),
                };

                rows.Add(row);
            }

            var appendRequest = new AppendCellsRequest
            {
                Fields = "*",
                SheetId = this.SheetID,
                Rows = rows
            };

            Request request = new Request
            {
                AppendCells = appendRequest
            };

            // Wrap it into batch update request.
            BatchUpdateSpreadsheetRequest batchRequst = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { request }
            };

            // Finally update the sheet.
            return this.Service.Spreadsheets
                .BatchUpdate(batchRequst, this.SpreadsheetID)
                .Execute();
        }
    }
}
