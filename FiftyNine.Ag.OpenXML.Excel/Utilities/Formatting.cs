using System;
using System.Net;
using System.Windows;
using System.Windows.Input;

namespace FiftyNine.Ag.OpenXML.Excel.Utilities
{
    public static class Formatting
    {
        public static string ColRowFormat(int col, int row)
        {
            return string.Format(string.Format("{0}{1}", (char)(col + 64), row + 1));
        }
    }
}
