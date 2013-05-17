using System;
using System.Net;
using System.Windows;
using System.Windows.Input;

namespace FiftyNine.Ag.OpenXML.Excel.Utilities
{
    public static class SpreadsheetRelationshipTypes
    {
        public const string Workbook = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const string Worksheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        public const string SharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        public const string Table = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
        public const string WorkbookStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    }
}
