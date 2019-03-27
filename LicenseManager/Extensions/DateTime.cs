using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HGM.Hotbird64.LicenseManager.Extensions
{
    public static class DateTimeExtensions
    {
        public static string ToEpidPart(this DateTime date) => $"{date.DayOfYear:D3}{date.Year:D4}";
    }
}
