using System;

namespace GoogleSheetsWrapper
{
    /// <summary>
    /// Configuration for batching rows when appending data to a Google Sheet.
    /// </summary>
    public class BatchInformation
    {
        /// <summary>
        /// The delay in milliseconds between batch requests.
        /// </summary>
        public int BatchWaitTime { get; set; } = 1000;

        /// <summary>
        /// The maximum number of rows in each batch request.
        /// </summary>
        public int BatchSize { get; set; } = 100;

        /// <summary>
        /// Indicates whether the header row should be excluded from the sheet.
        /// </summary>
        public bool SkipWritingHeaderRow { get; set; }

        /// <summary>
        /// Called before waiting between batches. Return true to stop appending additional records.
        /// </summary>
        public Func<bool> ShouldCancel { get; set; }
    }
}
