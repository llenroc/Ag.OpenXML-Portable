using System;
using System.Net;
using FiftyNine.Ag.OpenXML.Common;
using System.Collections.Generic;
using System.Xml;
using FiftyNine.Ag.OpenXML.Excel.Utilities;
using FiftyNine.Ag.OpenXML.Common.BaseClasses;

namespace FiftyNine.Ag.OpenXML.Excel.Elements
{
    public abstract class SpreadsheetElement : OpenXMLElement
    {
        public override void Save(XmlWriter writer)
        {
            writer.WriteStartElement(NodeName);
            SaveContent(writer);
            writer.WriteEndElement();
        }

        protected override Namespace Namespace
        {
            get { return SpreadsheetNamespaces.Main; }
        }
    }
}
