namespace GoogleSheetsWrapper
{
    /// <summary>
    /// Configuration object for BaseRepository classes
    /// </summary>
    public class BaseRepositoryConfiguration
    {
        /// <summary>
        /// Does the Google Sheet use a Header Row?
        /// </summary>
        public bool HasHeaderRow { get; set; }

        /// <summary>
        /// The number of rows to skip before reading the first header row.  
        /// Leave this value 0 if the first row is also the header row
        /// </summary>
        public int HeaderRowOffset { get; set; }

        /// <summary>
        /// The number of rows to skip before reading the data table.
        /// If the spreadsheet has a header row, this offset starts after the header row
        /// If the spreadsheet does not have a header row, this offset starts at the beginning of the spreadsheet.
        /// </summary>
        public int DataTableRowOffset { get; set; }
    }
}
