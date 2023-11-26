using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// Data Transfer Object for BatchUpdateRequests
    /// </summary>
    public class BatchUpdateRequestObject
    {
        /// <summary>
        /// The data range
        /// </summary>
        public SheetRange Range { get; set; }

        /// <summary>
        /// The actual data
        /// </summary>
        public CellData Data { get; set; }

        /// <summary>
        /// The <see cref="SheetFieldAttribute"/> value
        /// </summary>
        public SheetFieldAttribute FieldAttribute { get; set; }
    }
}