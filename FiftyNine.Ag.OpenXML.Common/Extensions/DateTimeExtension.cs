using System;
using System.Net;
using System.Xml;

namespace FiftyNine.Ag.OpenXML.Common.Extensions
{
    public static class DateTimeExtension
    {
        public static string ToXmlString(this DateTime dt)
        {
            return dt.ToString("yyyy-MM-ddTHH:mm:ssZ");
        }
    }
}
