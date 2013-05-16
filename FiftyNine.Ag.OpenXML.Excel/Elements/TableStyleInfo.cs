using System;
using System.Net;
using System.Windows;
using System.Windows.Input;
using FiftyNine.Ag.OpenXML.Common.Packaging;
using System.Xml;

namespace FiftyNine.Ag.OpenXML.Excel.Elements
{
    public class TableStyleInfo : SpreadsheetElement
    {
        protected override void SaveContent(XmlWriter writer)
        {
            writer.WriteAttributeString("showFirstColumn", ShowFirstColumn ? "1" : "0");
            writer.WriteAttributeString("showLastColumn", ShowLastColumn ? "1" : "0");
            writer.WriteAttributeString("showRowStripes", ShowRowStripes ? "1" : "0");
            writer.WriteAttributeString("showColumnStripes", ShowColumnStripes ? "1" : "0");
        }

        public bool ShowFirstColumn
        {
            get;
            set;
        }
        public bool ShowLastColumn
        {
            get;
            set;
        }
        public bool ShowRowStripes
        {
            get;
            set;
        }
        public bool ShowColumnStripes
        {
            get;
            set;
        }

        protected override string NodeName
        {
            get { return "tableStyleInfo"; }
        }
    }
}
