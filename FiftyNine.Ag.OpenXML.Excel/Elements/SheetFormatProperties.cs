using System;
using System.Net;
using System.Windows;
using System.Windows.Input;
using FiftyNine.Ag.OpenXML.Common;
using System.Collections.Generic;
using FiftyNine.Ag.OpenXML.Excel.Utilities;
using FiftyNine.Ag.OpenXML.Common.Packaging;

namespace FiftyNine.Ag.OpenXML.Excel.Elements
{
    public class SheetFormatProperties : SpreadsheetElement
    {
        protected override void Initialize()
        {
            DefaultRowHeight = 15;
        }

        protected override void SaveContent(System.Xml.XmlWriter writer)
        {
            writer.WriteAttributeString("defaultRowHeight", DefaultRowHeight.ToString());
        }

        public double DefaultRowHeight
        {
            get;
            set;
        }

        protected override string NodeName
        {
            get { return "sheetFormatPr"; }
        }
    }
}
