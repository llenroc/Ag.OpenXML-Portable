using System;
using System.Net;
using System.Windows;
using System.Windows.Input;

namespace FiftyNine.Ag.OpenXML.Excel.Utilities
{
    public static class SpreadsheetContentTypes
    {
        public const string Workbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        public const string Worksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        public const string SharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        public const string Table = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
        public const string WorkbookStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    }
}
