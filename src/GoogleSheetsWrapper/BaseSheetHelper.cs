using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace GoogleSheetsWrapper
{
    public abstract class BaseSheetHelper
    {
        public string SpreadsheetID { get; set; }

        public string TabName { get; set; }

        public int? SheetID { get; set; }

        public string ServiceAccountEmail { get; set; }

        public string[] Scopes { get; set; } = { SheetsService.Scope.Spreadsheets };

        public SheetsService Service { get; private set; }

        public BaseSheetHelper(string spreadsheetID, string serviceAccountEmail, string tabName)
        {
            this.SpreadsheetID = spreadsheetID;
            this.ServiceAccountEmail = serviceAccountEmail;
            this.TabName = tabName;
        }

        public void Init(string jsonCredentials)
        {
            var credential = (ServiceAccountCredential)
                   GoogleCredential.FromJson(jsonCredentials).UnderlyingCredential;

            // Authenticate as service account to the Sheets API
            var initializer = new ServiceAccountCredential.Initializer(credential.Id)
            {
                User = this.ServiceAccountEmail,
                Key = credential.Key,
                Scopes = Scopes
            };
            credential = new ServiceAccountCredential(initializer);

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
            });

            this.Service = service;

            this.UpdateTabName(this.TabName);
        }

        public void UpdateTabName(string newTabName)
        {
            // Lookup the sheet id for the given tab name
            if (!string.IsNullOrEmpty(newTabName))
            {
                var spreadsheet = this.Service.Spreadsheets.Get(this.SpreadsheetID);

                var result = spreadsheet.Execute();

                var sheet = result.Sheets.Where(s => s.Properties.Title.Equals(newTabName, StringComparison.CurrentCultureIgnoreCase)).First();

                this.SheetID = sheet.Properties.SheetId;
            }
            else
            {
                var spreadsheet = this.Service.Spreadsheets.Get(this.SpreadsheetID);

                var result = spreadsheet.Execute();

                var sheet = result.Sheets.First();

                this.SheetID = sheet.Properties.SheetId;
            }
        }

        public List<string> GetAllTabNames()
        {
            var spreadsheet = this.Service.Spreadsheets.Get(this.SpreadsheetID);

            var result = spreadsheet.Execute();

            var tabs = result.Sheets.Select(s => s.Properties.Title).ToList();

            return tabs;
        }

        public IList<IList<object>> GetRows(SheetRange range)
        {
            var rangeValue = range.CanSupportA1Notation ? range.A1Notation : range.R1C1Notation;

            GetRequest request =
                    this.Service.Spreadsheets.Values.Get(this.SpreadsheetID, rangeValue);

            request.ValueRenderOption = GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
            request.DateTimeRenderOption = GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;

            ValueRange response = request.Execute();
            return response.Values;
        }

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

        public BatchUpdateSpreadsheetResponse DeleteRow(int row)
        {
            var requests = new List<Request>();

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

            requests.Add(request);

            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest();
            bussr.Requests = requests;

            var updateRequest = this.Service.Spreadsheets.BatchUpdate(bussr, this.SpreadsheetID);
            return updateRequest.Execute();
        }

        public BatchUpdateSpreadsheetResponse BatchUpdate(List<BatchUpdateRequestObject> updates)
        {
            BatchUpdateSpreadsheetRequest bussr = new BatchUpdateSpreadsheetRequest();

            var requests = new List<Request>();

            foreach (var update in updates)
            {
                //create the update request for cells from the first row
                var updateCellsRequest = new Request()
                {
                    RepeatCell = new RepeatCellRequest()
                    {
                        Range = new GridRange()
                        {
                            SheetId = this.SheetID,
                            StartColumnIndex = update.Range.StartColumn - 1,
                            StartRowIndex = update.Range.StartRow - 1,
                            EndColumnIndex = update.Range.StartColumn,
                            EndRowIndex = update.Range.StartRow,
                        },
                        Cell = update.Data,
                        Fields = "*"
                    }
                };

                requests.Add(updateCellsRequest);
            }

            bussr.Requests = requests;

            var updateRequest = this.Service.Spreadsheets.BatchUpdate(bussr, this.SpreadsheetID);
            return updateRequest.Execute();
        }
    }
}
