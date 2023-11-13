using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheetsWrapper
{
    public class SchemaValidationResult
    {
        public bool IsValid { get; set; }

        public string ErrorMessage { get; set; }
    }
}