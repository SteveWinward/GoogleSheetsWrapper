using System;

namespace GoogleSheetsWrapper.Utils
{
    /// <summary>
    /// Utility class for parsing DateTime values
    /// </summary>
    public class DateTimeUtils
    {
        private static readonly DateTime EpochDate = new DateTime(1899, 12, 30);

        private static readonly double SecondsInADay = 86400.0;

        /// <summary>
        /// Convert to Serial Number format
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static double ConvertToSerialNumber(DateTime dateTime)
        {
            var t = dateTime - EpochDate;

            var secondsSinceEpoch = t.TotalSeconds;

            return secondsSinceEpoch / SecondsInADay;
        }

        /// <summary>
        /// Convert from Serial Number format
        /// </summary>
        /// <param name="serialNumber"></param>
        /// <returns></returns>
        public static DateTime ConvertFromSerialNumber(double serialNumber)
        {
            var totalSeconds = serialNumber * SecondsInADay;

            return EpochDate.AddSeconds(totalSeconds);
        }

        /// <summary>
        /// Returns current EST time
        /// </summary>
        /// <returns></returns>
        public static DateTime GetCurrentEstTime()
        {
            var timeUtc = DateTime.UtcNow;
            var easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            return TimeZoneInfo.ConvertTimeFromUtc(timeUtc, easternZone);
        }
    }
}