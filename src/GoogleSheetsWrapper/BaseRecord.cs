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

        /// <summary>
        /// Convert a Google Sheet API row result into a BaseRecord
        /// </summary>
        /// <param name="row"></param>
        /// <param name="rowId"></param>
        public BaseRecord(IList<object> row, int rowId)
            : this(row, rowId, 1)
        {
        }

        /// <summary>
        /// Convert a Google Sheet API row result into a BaseRecord with an column id offset
        /// </summary>
        /// <param name="row"></param>
        /// <param name="rowId"></param>
        /// <param name="minColumnId"></param>
        public BaseRecord(IList<object> row, int rowId, int minColumnId)
        {
            this.RowId = rowId;
            SheetFieldAttributeUtils.PopulateRecord(this, row, minColumnId);
        }

        /// <summary>
        /// Converts this current object into the Google Sheet's associated CellData format required for the Google Sheets API
        /// </summary>
        /// <param name="tabName"></param>
        /// <returns></returns>
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
