using System;
using System.Net;
using System.Windows;
using System.Windows.Input;

namespace FiftyNine.Ag.OpenXML.Excel.Elements
{
    public class CellDefinition
    {
        public CellDefinition(string col, string row)
        {
            Column = col;
            Row = row;
        }

        public string Column
        {
            get;
            private set;
        }
        public string Row
        {
            get;
            private set;
        }
    }
}
