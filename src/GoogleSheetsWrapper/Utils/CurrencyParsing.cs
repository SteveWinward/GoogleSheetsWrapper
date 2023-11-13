using System.Globalization;

namespace GoogleSheetsWrapper.Utils
{
    /// <summary>
    /// Utility class for parsing currency values
    /// </summary>
    public class CurrencyParsing
    {
        /// <summary>
        /// Parse currency text value to double
        /// </summary>
        /// <param name="amount"></param>
        /// <returns></returns>
        public static double ParseCurrencyString(string amount)
        {
            return double.Parse(
                amount,
                NumberStyles.Currency | NumberStyles.AllowDecimalPoint, CultureInfo.CreateSpecificCulture("en-us"));
        }

        /// <summary>
        /// Convert currency double value to string currency format
        /// </summary>
        /// <param name="amount"></param>
        /// <returns></returns>
        public static string ConvertToCurrencyString(double amount)
        {
            return $"{amount:C2}";
        }
    }
}