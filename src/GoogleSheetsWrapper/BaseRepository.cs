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
        }

        protected void InitSheetRanges()
        {
            var attributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes<T>();

            var minColumnId = attributes.Min(a => a.Key.ColumnID);
            var maxColumnId = attributes.Max(a => a.Key.ColumnID);

            this.SheetHeaderRange = new SheetRange(this.SheetsHelper.TabName, minColumnId, 1, maxColumnId, 1);
            this.SheetDataRange = new SheetRange(this.SheetsHelper.TabName, minColumnId, 2, maxColumnId);
        }

        public void AddRecord(T record)
        {
            this.SheetsHelper.AppendRow(record);
        }

        public void AddRecords(List<T> records)
        {
            this.SheetsHelper.AppendRows(records);
        }

        public void DeleteRecord(T record)
        {
            this.SheetsHelper.DeleteRow(record.RowId);
        }

        public List<BatchUpdateRequestObject> FilterUpdates(
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

        public BatchUpdateSpreadsheetResponse SaveFields(T record, List<Expression<Func<T, object>>> properties)
        {
            var data = record.ConvertToCellData(this.SheetsHelper.TabName);

            var updates = this.FilterUpdates(data, properties);

            return this.SheetsHelper.BatchUpdate(updates);
        }

        public BatchUpdateSpreadsheetResponse SaveField(T record, Expression<Func<T, object>> expression)
        {
            return this.SaveFields(record, new List<Expression<Func<T, object>>>() { expression });
        }

        public SchemaValidationResult ValidateSchema()
        {
            var headers = this.SheetsHelper.GetRows(SheetHeaderRange);

            return this.ValidateSchema(headers[0]);
        }

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
    }
}