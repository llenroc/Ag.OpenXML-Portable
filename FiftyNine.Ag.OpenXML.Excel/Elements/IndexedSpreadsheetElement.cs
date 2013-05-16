using System;
using System.Net;
using System.Windows;
using System.Windows.Input;
using FiftyNine.Ag.OpenXML.Common.Packaging;

namespace FiftyNine.Ag.OpenXML.Excel.Elements
{
    public abstract class IndexedSpreadsheetElement : SpreadsheetElement
    {
        internal void SetUp(int index)
        {
            Index = index;
        }

        protected int Index
        {
            get;
            private set;
        }
    }
}
