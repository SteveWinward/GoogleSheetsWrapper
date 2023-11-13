namespace GoogleSheetsWrapper
{

#pragma warning disable CA1720 // String / Integer contains type name

    /// <summary>
    /// SheetFieldType enum class
    /// </summary>
    public enum SheetFieldType
    {
        /// <summary>
        /// A string data value
        /// </summary>
        String,
        /// <summary>
        /// A DateTime data vlue
        /// </summary>
        DateTime,
        /// <summary>
        /// A Currency data value
        /// </summary>
        Currency,
        /// <summary>
        /// A PhoneNumber data value
        /// </summary>
        PhoneNumber,
        /// <summary>
        /// A Boolean data value
        /// </summary>
        Boolean,
        /// <summary>
        /// An Integer data value (ie int)
        /// </summary>
        Integer,
        /// <summary>
        /// A Number data value (ie double)
        /// </summary>
        Number
    }

#pragma warning restore CA1720 // String / Integer contains type name

}