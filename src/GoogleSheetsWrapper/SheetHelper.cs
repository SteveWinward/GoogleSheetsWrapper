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

namespace GoogleSheetsWrapper
{
    public class SheetHelper<T> : BaseSheetHelper where T : BaseRecord
    {
        public SheetHelper(string spreadsheetID, string serviceAccountEmail, string tabName)
            : base(spreadsheetID, serviceAccountEmail, tabName) { }

        public void AppendRow(T record)
        {
            this.AppendRows(new List<T>() { record });
        }

        public void AppendRows(List<T> records)
        {
            var rows = new List<RowData>();

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
            this.Service.Spreadsheets
                .BatchUpdate(batchRequst, this.SpreadsheetID)
                .Execute();
        }
    }
}
