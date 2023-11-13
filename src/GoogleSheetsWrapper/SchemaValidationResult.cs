namespace GoogleSheetsWrapper
{
    /// <summary>
    /// Simple result object for schema validation on Google Sheets tabs
    /// </summary>
    public class SchemaValidationResult
    {
        /// <summary>
        /// Is the schema valid?
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// Error message if the validation fails
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}