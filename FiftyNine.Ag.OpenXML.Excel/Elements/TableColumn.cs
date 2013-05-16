using System;
using System.Net;
using System.Windows;
using System.Windows.Input;
using FiftyNine.Ag.OpenXML.Common.Packaging;

namespace FiftyNine.Ag.OpenXML.Excel.Elements
{
    public class TableColumn : SpreadsheetElement
    {
        protected override void SaveContent(System.Xml.XmlWriter writer)
        {
            writer.WriteAttributeString("id", ID.ToString());
            writer.WriteAttributeString("name", Name);
        }

        public int ID
        {
            get;
            internal set;
        }
        public string Name
        {
            get;
            set;
        }

        protected override string NodeName
        {
            get { return "tableColumn"; }
        }
    }
}
