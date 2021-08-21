using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;

namespace GoogleSheetsWrapper
{
    public abstract class BaseRepository<T> where T : BaseRecord
    {
        protected SheetHelper<T> SheetsHelper { get; set; }

        public SheetRange SheetDataRange { get; private set; }

        public SheetRange SheetHeaderRange { get; private set; }

        public BaseRepository() { }

        public BaseRepository(SheetHelper<T> sheetsHelper)
        {
            this.SheetsHelper = sheetsHelper;

            this.InitSheetRanges();
        }

        public BaseRepository(string spreadsheetID, string serviceAccountEmail, string tabName, string jsonCredentials)
        {
            this.SheetsHelper = new SheetHelper<T>(spreadsheetID, serviceAccountEmail, tabName);

            this.SheetsHelper.Init(jsonCredentials);

            this.InitSheetRanges();
        }

        /// <summary>
        /// Determines the proper sheet ranges to query the Sheets table
        /// </summary>
        protected void InitSheetRanges()
        {
            var attributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes<T>();

            var maxColumnId = attributes.Max(a => a.Key.ColumnID);

            this.SheetHeaderRange = new SheetRange(this.SheetsHelper.TabName, 1, 1, maxColumnId, 1);
            this.SheetDataRange = new SheetRange(this.SheetsHelper.TabName, 1, 2, maxColumnId);
        }

        /// <summary>
        /// Create a new record to the end of the current sheets table
        /// </summary>
        /// <param name="record"></param>
        public BatchUpdateSpreadsheetResponse AddRecord(T record)
        {
            return this.SheetsHelper.AppendRow(record);
        }

        /// <summary>
        /// Create a list of records to the end of the current sheets table
        /// </summary>
        /// <param name="records"></param>
        public BatchUpdateSpreadsheetResponse AddRecords(List<T> records)
        {
            return this.SheetsHelper.AppendRows(records);
        }

        /// <summary>
        /// Delete the specified record
        /// </summary>
        /// <param name="record"></param>
        public void DeleteRecord(T record)
        {
            this.SheetsHelper.DeleteRow(record.RowId);
        }

        /// <summary>
        /// Get all of the records back from the Google Sheets tab
        /// </summary>
        /// <returns></returns>
        public List<T> GetAllRecords()
        {
            var result = this.SheetsHelper.GetRows(SheetDataRange);

            var records = new List<T>();

            for (int r = 0; r < result.Count; r++)
            {
                var currentRecord = result[r];

                var record = (T)Activator.CreateInstance(typeof(T), currentRecord, r + this.SheetDataRange.StartRow);
                records.Add(record);
            }

            return records;
        }

        /// <summary>
        /// Save multiple field values to the row
        /// </summary>
        /// <param name="record"></param>
        /// <param name="properties"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse SaveFields(T record, List<Expression<Func<T, object>>> properties)
        {
            var data = record.ConvertToCellData(this.SheetsHelper.TabName);

            var updates = this.FilterUpdates(data, properties);

            return this.SheetsHelper.BatchUpdate(updates);
        }

        /// <summary>
        /// Save a single field value to the row
        /// </summary>
        /// <param name="record"></param>
        /// <param name="expression"></param>
        /// <returns></returns>
        public BatchUpdateSpreadsheetResponse SaveField(T record, Expression<Func<T, object>> expression)
        {
            return this.SaveFields(record, new List<Expression<Func<T, object>>>() { expression });
        }

        /// <summary>
        /// Ensure the header names for the Google Sheet matches the expected names based on the SheetFieldAttribute metadata
        /// </summary>
        /// <returns></returns>
        public SchemaValidationResult ValidateSchema()
        {
            var headers = this.SheetsHelper.GetRows(SheetHeaderRange);

            return this.ValidateSchema(headers[0]);
        }

        /// <summary>
        /// Ensure the header names for the Google Sheet matches the expected names based on the SheetFieldAttribute metadata
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public SchemaValidationResult ValidateSchema(IList<object> row)
        {
            var result = new SchemaValidationResult()
            {
                IsValid = true,
                ErrorMessage = "",
            };

            var fieldAttributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes<T>();

            foreach (var kvp in fieldAttributes)
            {
                var attribute = kvp.Key;

                if (row.Count < (attribute.ColumnID))
                {
                    result.IsValid = false;
                    result.ErrorMessage += $"'{attribute.DisplayName}' column id: {attribute.ColumnID} is greater than the defined column count: {row.Count}. ";
                }

                else if (row[attribute.ColumnID - 1]?.ToString() != attribute.DisplayName)
                {
                    result.IsValid = false;
                    result.ErrorMessage += $"Expected column name: '{attribute.DisplayName}' however, current column name is: '{row[attribute.ColumnID - 1]}'. ";
                }
            }

            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="expressions"></param>
        /// <returns></returns>
        internal List<BatchUpdateRequestObject> FilterUpdates(
            List<BatchUpdateRequestObject> data,
            List<Expression<Func<T, object>>> expressions)
        {
            var result = new List<BatchUpdateRequestObject>();

            foreach (var exp in expressions)
            {
                var columnID = SheetFieldAttributeUtils.GetColumnId(exp);
                result.Add(data.Where(b => b.FieldAttribute.ColumnID == columnID).First());
            }

            return result;
        }
    }
}