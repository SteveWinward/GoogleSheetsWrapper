using System.Collections.Generic;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// BaseRecord abstract class
    /// </summary>
    public abstract class BaseRecord
    {
        /// <summary>
        /// The Row ID in the Google Sheets tab
        /// </summary>
        public int RowId { get; set; }

        /// <summary>
        /// Default constructor for BaseRecord
        /// </summary>
        public BaseRecord() { }

        /// <summary>
        /// Convert a Google Sheet API row result into a BaseRecord with an column id offset
        /// </summary>
        /// <param name="row"></param>
        /// <param name="rowId"></param>
        /// <param name="minColumnId"></param>
        public BaseRecord(IList<object> row, int rowId, int minColumnId = 1)
        {
            RowId = rowId;
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

            var attributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes(GetType());

            foreach (var attribute in attributes)
            {
                var sheetRange = new SheetRange(tabName, attribute.Key.ColumnID, RowId);

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