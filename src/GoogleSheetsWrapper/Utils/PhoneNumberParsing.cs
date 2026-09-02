using System.Globalization;
using System.Text.RegularExpressions;

namespace GoogleSheetsWrapper.Utils
{
    /// <summary>
    /// Utility class for parsing Phone Numbers
    /// </summary>
    public class PhoneNumberParsing
    {
        /// <summary>
        /// Removes any occurrence of the United States country code ("+1") and all non-numeric characters.
        /// </summary>
        /// <param name="number">The phone number to normalize.</param>
        /// <returns>The normalized phone number containing digits only.</returns>
        public static string RemoveUSInterationalPhoneCode(string number)
        {
            return RemoveExtraCharactersFromPhoneNumber(number.Replace("+1", ""));
        }

        /// <summary>
        /// Removes all non-numeric characters from a phone number.
        /// </summary>
        /// <param name="number">The phone number to normalize.</param>
        /// <returns>The phone number containing digits only.</returns>
        public static string RemoveExtraCharactersFromPhoneNumber(string number)
        {
            return Regex.Replace(number, @"[^\d]", "");
        }

        /// <summary>
        /// Converts a normalized United States phone number to a numeric value.
        /// </summary>
        /// <param name="number">The phone number to convert.</param>
        /// <returns>The numeric phone number with any occurrence of the United States country code ("+1") removed.</returns>
        public static long ConvertToLong(string number)
        {
            var numberAsString = RemoveExtraCharactersFromPhoneNumber(number.Replace("+1", ""));

            return long.Parse(numberAsString, CultureInfo.CurrentCulture);
        }
    }
}