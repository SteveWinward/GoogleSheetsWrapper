using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheetsWrapper.Utils
{
    public class DateTimeUtils
    {
        private static DateTime EpochDate = new DateTime(1899, 12, 30);

        private static double SecondsInADay = 86400.0;

        public static double ConvertToSerialNumber(DateTime dateTime)
        {
            TimeSpan t = dateTime - EpochDate;

            var secondsSinceEpoch = t.TotalSeconds;

            return secondsSinceEpoch / SecondsInADay;
        }

        public static DateTime ConvertFromSerialNumber(double serialNumber)
        {
            var totalSeconds = serialNumber * SecondsInADay;

            return EpochDate.AddSeconds(totalSeconds);
        }

        public static DateTime GetCurrentEstTime()
        {
            var timeUtc = DateTime.UtcNow;
            TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            return TimeZoneInfo.ConvertTimeFromUtc(timeUtc, easternZone);
        }
    }
}
