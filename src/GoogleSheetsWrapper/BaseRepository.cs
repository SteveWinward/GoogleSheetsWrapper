using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// BaseRepository abstract base class
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class BaseRepository<T> where T : BaseRecord
    {
        /// <summary>
        /// SheetHelper class associated with the BaseRepository
        /// </summary>
        protected SheetHelper<T> SheetHelper { get; set; }

        /// <summary>
        /// The SheetRange object representing the data range in the Google Sheets tab (does not include the headers if there are headers)
        /// </summary>
        public SheetRange SheetDataRange { get; private set; }

        /// <summary>
        /// The SheetRange object representing the header range in the Google Sheets tab
        /// </summary>
        public SheetRange SheetHeaderRange { get; private set; }

        /// <summary>
        /// Does the tab have headers?
        /// </summary>
        public bool HasHeaderRow { get; private set; }

        private bool BaseRecordConstructorWasVerified;

        private bool BaseRecordConstructorExists;

        private readonly BaseRepositoryConfiguration Configuration;

        /// <summary>
        /// Default constructor
        /// </summary>
        public BaseRepository() { }

        /// <summary>
        /// BaseRepository constructor
        /// </summary>
        /// <param name="sheetHelper">SheetHelper object</param>
        /// <param name="config">BaseRepositoryConfiguration config options for the repository class</param>
        public BaseRepository(SheetHelper<T> sheetHelper, BaseRepositoryConfiguration config)
        {
            SheetHelper = sheetHelper;

            Configuration = config;

            InitSheetRanges();
        }

        /// <summary>
        /// BaseRepository constructor
        /// </summary>
        /// <param name="sheetHelper">SheetHelper object</param>
        /// <param name="hasHeaderRow">Does the tab have headers?</param>
        public BaseRepository(SheetHelper<T> sheetHelper, bool hasHeaderRow = true)
            : this(sheetHelper, new BaseRepositoryConfiguration() { HasHeaderRow = hasHeaderRow })
        {

        }

        /// <summary>
        /// BaseRepository constructor
        /// </summary>
        /// <param name="spreadsheetID">The Google Sheets spreadsheet ID</param>
        /// <param name="serviceAccountEmail">The Google Sheets service account email</param>
        /// <param name="tabName">The Google Sheets tab name</param>
        /// <param name="jsonCredentials">The Google Sheets JSON credentials</param>
        /// <param name="hasHeaderRow">Does the tab have headers?</param>
        [Obsolete("This constructer is deprecated.  Please use the BaseRepository(SheetHelper<T> sheetHelper, BaseRepositoryConfiguration config) moving forward.")]
        public BaseRepository(string spreadsheetID, string serviceAccountEmail, string tabName, string jsonCredentials, bool hasHeaderRow = true)
        {
            SheetHelper = new SheetHelper<T>(spreadsheetID, serviceAccountEmail, tabName);

            SheetHelper.Init(jsonCredentials);

            Configuration = new BaseRepositoryConfiguration()
            {
                HasHeaderRow = hasHeaderRow,
            };

            InitSheetRanges();
        }

        /// <summary>
        /// Determines the proper sheet ranges to query the Sheets table
        /// </summary>
        protected void InitSheetRanges()
        {
            HasHeaderRow = Configuration.HasHeaderRow;

            var attributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes<T>();

            var maxColumnId = attributes.Max(a => a.Key.ColumnID);
            var minColumnId = attributes.Min(a => a.Key.ColumnID);

            if (HasHeaderRow)
            {
                var headerRow = 1 + Configuration.HeaderRowOffset;
                var dataTableRowStart = headerRow + 1 + Configuration.DataTableRowOffset;

                SheetHeaderRange = new SheetRange(SheetHelper.TabName, minColumnId, headerRow, maxColumnId, headerRow);
                SheetDataRange = new SheetRange(SheetHelper.TabName, minColumnId, dataTableRowStart, maxColumnId);
            }
            else
            {
                var dataTableRowStart = 1 + Configuration.DataTableRowOffset;

                SheetDataRange = new SheetRange(SheetHelper.TabName, minColumnId, dataTableRowStart, maxColumnId);
            }
        }

        /// <summary>
        /// Create a new record to the end of the current sheets table
        /// </summary>
        /// <param name="record"></param>
        public BatchUpdateSpreadsheetResponse AddRecord(T record)
        {
            return SheetHelper.AppendRow(record);
        }

        /// <summary>
        /// Create a list of records to the end of the current sheets table
        /// </summary>
        /// <param name="records"></param>
        public BatchUpdateSpreadsheetResponse AddRecords(IList<T> records)
        {
            return SheetHelper.AppendRows(records);
        }

        /// <summary>
        /// Delete the specified record
        /// </summary>
        /// <param name="record"></param>
        public void DeleteRecord(T record)
        {
            _ = SheetHelper.DeleteRow(record.RowId);
        }

        /// <summary>
        /// Get all of the records back from the Google Sheets tab
        /// </summary>
        /// <returns></returns>
        public List<T> GetAllRecords()
        {
            var result = SheetHelper.GetRows(SheetDataRange);

            List<T> records;

            // Check if the result is null first
            if (result == null)
            {
                records = new List<T>();
            }
            // Otherwise, loop over all rows to create strongly typed objects for each row
            else
            {
                // Validate the BaseRecord class has the correct constructor signature
                if (!ValidateBaseRecordHasConstructorDefined())
                {
                    throw new ArgumentException(
                        $@"Type {typeof(T).Name} does not implement a constructor with parameters 
                        (IList<object> row, int rowId, int minColumnId).  
                        Please define this constructor signature and recompile your code");
                }

                records = new List<T>(result.Count);

                for (var r = 0; r < result.Count; r++)
                {
                    var currentRecord = result[r];

                    var record = (T)Activator.CreateInstance(
                        typeof(T),
                        currentRecord,
                        r + SheetDataRange.StartRow,
                        SheetDataRange.StartColumn);

                    records.Add(record);
                }
            }

            return records;
        }

        /// <summary>
        /// Save a single field value to the row
        /// </summary>
        public BatchUpdateSpreadsheetResponse SaveField(T record, Expression<Func<T, object>> expression)
        {
            return SaveFields(record, new Expression<Func<T, object>>[] { expression } as IList<Expression<Func<T, object>>>);
        }

        /// <summary>
        /// Save multiple field values to the row. Example: SaveFields(record, (r) => r.Property1, (r) => r.Property2);
        /// </summary>
        public BatchUpdateSpreadsheetResponse SaveFields(T record, params Expression<Func<T, object>>[] expressions)
        {
            return SaveFields(record, expressions as IList<Expression<Func<T, object>>>); //casts prevent stackoverflow to params
        }

        /// <summary>
        /// Save multiple field values to the row
        /// </summary>
        public BatchUpdateSpreadsheetResponse SaveFields(T record, IList<Expression<Func<T, object>>> properties)
        {
            var data = record.ConvertToCellData(SheetHelper.TabName);

            var updates = FilterUpdates(data, properties);

            return SheetHelper.BatchUpdate(updates);
        }

        /// <summary>
        /// Ensure the header names for the Google Sheet matches the expected names based on the SheetFieldAttribute metadata
        /// </summary>
        /// <returns></returns>
        public SchemaValidationResult ValidateSchema()
        {
            if (!HasHeaderRow)
            {
                throw new ArgumentException("ValidateSchema cannot be called when the HasHeaderRow property is set to false");
            }

            var headers = SheetHelper.GetRows(SheetHeaderRange);

            return ValidateSchema(headers[0]);
        }

        /// <summary>
        /// Ensure the header names for the Google Sheet matches the expected names based on the SheetFieldAttribute metadata
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public SchemaValidationResult ValidateSchema(IList<object> row)
        {
            if (!HasHeaderRow)
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

                if (row.Count < attribute.ColumnID)
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
            if (HasHeaderRow)
            {
                throw new ArgumentException("HasHeaderRow must be false to call this method");
            }

            var fieldAttributes = SheetFieldAttributeUtils.GetAllSheetFieldAttributes<T>();
            var fieldNames = new List<string>();

            foreach (var kvp in fieldAttributes)
            {
                fieldNames.Add(kvp.Key.DisplayName);
            }

            var appender = new SheetAppender(SheetHelper);

            appender.AppendRow(fieldNames);
        }

        /// <summary>
        /// Filters the BatchUpdateRequest to only the actual fields specified in the list
        /// </summary>
        /// <param name="data"></param>
        /// <param name="expressions"></param>
        /// <returns></returns>
        internal static List<BatchUpdateRequestObject> FilterUpdates(
            IList<BatchUpdateRequestObject> data,
            IList<Expression<Func<T, object>>> expressions)
        {
            var result = new List<BatchUpdateRequestObject>();

            foreach (var exp in expressions)
            {
                var columnID = SheetFieldAttributeUtils.GetColumnId(exp);
                result.Add(data.First(b => b.FieldAttribute.ColumnID == columnID));
            }

            return result;
        }

        private bool ValidateBaseRecordHasConstructorDefined()
        {
            if (!BaseRecordConstructorWasVerified)
            {
                var types = new Type[]
                    {
                        typeof(IList<object>),
                        typeof(int),
                        typeof(int)
                    };

                var constructor = typeof(T).GetConstructor(types);

                BaseRecordConstructorExists = constructor != null;
                BaseRecordConstructorWasVerified = true;
            }
            return BaseRecordConstructorExists;
        }
    }
}