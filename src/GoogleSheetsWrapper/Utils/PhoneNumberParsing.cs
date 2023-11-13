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
        /// 
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public static string RemoveUSInterationalPhoneCode(string number)
        {
            return RemoveExtraCharactersFromPhoneNumber(number.Replace("+1", ""));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public static string RemoveExtraCharactersFromPhoneNumber(string number)
        {
            return Regex.Replace(number, @"[^\d]", "");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public static long ConvertToLong(string number)
        {
            var numberAsString = RemoveExtraCharactersFromPhoneNumber(number.Replace("+1", ""));

            return long.Parse(numberAsString, CultureInfo.CurrentCulture);
        }
    }
}