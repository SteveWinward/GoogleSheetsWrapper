using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace GoogleSheetsWrapper
{
    public class SheetExporter : BaseSheetHelper
    {
        public SheetExporter(string spreadsheetID, string serviceAccountEmail, string tabName) : base(spreadsheetID, serviceAccountEmail, tabName)
        {
        }

        public void ExportAsCsv(SheetRange range, StreamWriter streamWriter)
        {
            var rows = this.GetRowsFormatted(range);

            using (var csv = new CsvWriter(streamWriter, CultureInfo.InvariantCulture))
            {
                foreach (var row in rows)
                {
                    foreach (var cell in row)
                    {
                        csv.WriteField(cell?.ToString());
                    }

                    csv.NextRecord();
                }
            }
        }
    }
}
