using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class DateTimeHelper
    {
        public static string To8DigitString(this DateTime dtm) => dtm.ToString("yyyyMMdd");
        public static bool IsAfter(this DateTime dateTime, int year, int month, int day) => dateTime > new DateTime(year, month, day);
        public static bool IsOrAfter(this DateTime dateTime, int year, int month, int day) => dateTime >= new DateTime(year, month, day);
        public static bool IsTodayAfter(int year, int month, int day) => DateTime.Today > new DateTime(year, month, day);
        public static bool IsTodayAfter(int year, int month, int day, int minusDays) => DateTime.Today > new DateTime(year, month, day).AddDays(-minusDays);
        public static bool IsTodayOrAfter(int year, int month, int day) => DateTime.Today >= new DateTime(year, month, day);
        public static bool IsTodayOrAfter(int year, int month, int day, int minusDays) => DateTime.Today >= new DateTime(year, month, day).AddDays(-minusDays);
        public static bool IsBefore(this DateTime dateTime, int year, int month, int day) => dateTime < new DateTime(year, month, day);
        public static bool IsOrBefore(this DateTime dateTime, int year, int month, int day) => dateTime <= new DateTime(year, month, day);
        public static bool IsTodayBefore(int year, int month, int day) => DateTime.Today < new DateTime(year, month, day);
        public static bool IsTodayBefore(int year, int month, int day, int plusDays) => DateTime.Today < new DateTime(year, month, day).AddDays(plusDays);
        public static bool IsTodayOrBefore(int year, int month, int day) => DateTime.Today <= new DateTime(year, month, day);
        public static bool IsTodayOrBefore(int year, int month, int day, int plusDays) => DateTime.Today <= new DateTime(year, month, day).AddDays(plusDays);
        public static bool IsBetween(this DateTime dateTime, int year, int month, int day, int year2, int month2, int day2) => dateTime <= new DateTime(year2, month2, day2) && DateTime.Today >= new DateTime(year, month, day);
        public static bool IsTodayBetween(int year, int month, int day, int year2, int month2, int day2) => DateTime.Today <= new DateTime(year2, month2, day2) && DateTime.Today >= new DateTime(year, month, day);
    }
}
