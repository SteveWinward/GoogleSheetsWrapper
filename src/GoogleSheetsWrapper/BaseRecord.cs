using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace GoogleSheetsWrapper
{
    public abstract class BaseRecord
    {
        public int RowId { get; set; }

        public BaseRecord() { }

        public BaseRecord(IList<object> row, int rowId)
        {
            this.RowId = rowId;
            SheetFieldAttributeUtils.PopulateRecord(this, row);
        }

        public List<BatchUpdateRequestObject> ConvertToCellData(string tabName)
        {
            var results = new List<BatchUpdateRequestObject>();

            var attributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes(this.GetType());

            foreach (var attribute in attributes)
            {
                var sheetRange = new SheetRange(tabName, attribute.Key.ColumnID, this.RowId);

                results.Add(new BatchUpdateRequestObject()
                {
                    Range = sheetRange,
                    Data = SheetFieldAttributeUtils.GetCellDataForSheetField(this, attribute.Key, attribute.Value),
                    FieldAttribute = attribute.Key,
                });
            }

            return results;
        }
    }
}
