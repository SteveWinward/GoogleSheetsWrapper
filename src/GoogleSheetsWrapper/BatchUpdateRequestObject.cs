using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheetsWrapper
{
    public class BatchUpdateRequestObject
    {
        public SheetRange Range { get; set; }

        public CellData Data { get; set; }

        public SheetFieldAttribute FieldAttribute { get; set; }
    }
}
