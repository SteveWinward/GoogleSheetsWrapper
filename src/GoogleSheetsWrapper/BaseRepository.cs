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
        protected SheetHelper<T> SheetHelper { get; set; }

        public SheetRange SheetDataRange { get; private set; }

        public SheetRange SheetHeaderRange { get; private set; }

        public bool HasHeaderRow { get; private set; }

        public BaseRepository() { }

        public BaseRepository(SheetHelper<T> sheetHelper, bool hasHeaderRow = true)
        {
            this.SheetHelper = sheetHelper;

            this.HasHeaderRow = hasHeaderRow;

            this.InitSheetRanges();
        }

        public BaseRepository(string spreadsheetID, string serviceAccountEmail, string tabName, string jsonCredentials, bool hasHeaderRow = true)
        {
            this.SheetHelper = new SheetHelper<T>(spreadsheetID, serviceAccountEmail, tabName);

            this.SheetHelper.Init(jsonCredentials);

            this.HasHeaderRow = hasHeaderRow;

            this.InitSheetRanges();
        }

        /// <summary>
        /// Determines the proper sheet ranges to query the Sheets table
        /// </summary>
        protected void InitSheetRanges()
        {
            var attributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes<T>();

            var maxColumnId = attributes.Max(a => a.Key.ColumnID);
            var minColumnId = attributes.Min(a => a.Key.ColumnID);

            if (this.HasHeaderRow)
            {
                this.SheetHeaderRange = new SheetRange(this.SheetHelper.TabName, minColumnId, 1, maxColumnId, 1);
                this.SheetDataRange = new SheetRange(this.SheetHelper.TabName, minColumnId, 2, maxColumnId);
            }
            else
            {
                this.SheetDataRange = new SheetRange(this.SheetHelper.TabName, minColumnId, 1, maxColumnId);
            }
        }

        /// <summary>
        /// Create a new record to the end of the current sheets table
        /// </summary>
        /// <param name="record"></param>
        public BatchUpdateSpreadsheetResponse AddRecord(T record)
        {
            return this.SheetHelper.AppendRow(record);
        }

        /// <summary>
        /// Create a list of records to the end of the current sheets table
        /// </summary>
        /// <param name="records"></param>
        public BatchUpdateSpreadsheetResponse AddRecords(IList<T> records)
        {
            return this.SheetHelper.AppendRows(records);
        }

        /// <summary>
        /// Delete the specified record
        /// </summary>
        /// <param name="record"></param>
        public void DeleteRecord(T record)
        {
            this.SheetHelper.DeleteRow(record.RowId);
        }

        /// <summary>
        /// Get all of the records back from the Google Sheets tab
        /// </summary>
        /// <returns></returns>
        public List<T> GetAllRecords()
        {
            var result = this.SheetHelper.GetRows(SheetDataRange);

            var records = new List<T>(result.Count);

            for (int r = 0; r < result.Count; r++)
            {
                var currentRecord = result[r];

                var record = (T)Activator.CreateInstance(
                    typeof(T),
                    currentRecord,
                    r + this.SheetDataRange.StartRow,
                    this.SheetDataRange.StartColumn);

                records.Add(record);
            }

            return records;
        }

        /// <summary>
        /// Save a single field value to the row
        /// </summary>
        public BatchUpdateSpreadsheetResponse SaveField(T record, Expression<Func<T, object>> expression)
        {
            return this.SaveFields(record, new Expression<Func<T, object>>[] { expression } as IList<Expression<Func<T, object>>>);
        }

        /// <summary>
        /// Save multiple field values to the row. Example: SaveFields(record, (r) => r.Property1, (r) => r.Property2);
        /// </summary>
        public BatchUpdateSpreadsheetResponse SaveFields(T record, params Expression<Func<T, object>>[] expressions)
        {
            return this.SaveFields(record, expressions as IList<Expression<Func<T, object>>>); //casts prevent stackoverflow to params
        }

        /// <summary>
        /// Save multiple field values to the row
        /// </summary>
        public BatchUpdateSpreadsheetResponse SaveFields(T record, IList<Expression<Func<T, object>>> properties)
        {
            var data = record.ConvertToCellData(this.SheetHelper.TabName);

            var updates = this.FilterUpdates(data, properties);

            return this.SheetHelper.BatchUpdate(updates);
        }

        /// <summary>
        /// Ensure the header names for the Google Sheet matches the expected names based on the SheetFieldAttribute metadata
        /// </summary>
        /// <returns></returns>
        public SchemaValidationResult ValidateSchema()
        {
            if (!this.HasHeaderRow)
            {
                throw new ArgumentException("ValidateSchema cannot be called when the HasHeaderRow property is set to false");
            }

            var headers = this.SheetHelper.GetRows(SheetHeaderRange);

            return this.ValidateSchema(headers[0]);
        }

        /// <summary>
        /// Ensure the header names for the Google Sheet matches the expected names based on the SheetFieldAttribute metadata
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public SchemaValidationResult ValidateSchema(IList<object> row)
        {
            if (!this.HasHeaderRow)
            {
                throw new ArgumentException("ValidateSchema cannot be called when the HasHeaderRow property is set to false");
            }

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
        /// Writes the header names for a new tab
        /// 
        /// Only call this if the HasHeaderRow is set to false
        /// </summary>
        public void WriteHeaders()
        {
            if (this.HasHeaderRow)
            {
                throw new ArgumentException("HasHeaderRow must be false to call this method");
            }

            var fieldAttributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes<T>();
            var fieldNames = new List<string>();

            foreach (var kvp in fieldAttributes)
            {
                fieldNames.Add(kvp.Key.DisplayName);
            }

            var appender = new SheetAppender(this.SheetHelper);

            appender.AppendRow(fieldNames);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="expressions"></param>
        /// <returns></returns>
        internal List<BatchUpdateRequestObject> FilterUpdates(
            IList<BatchUpdateRequestObject> data,
            IList<Expression<Func<T, object>>> expressions)
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