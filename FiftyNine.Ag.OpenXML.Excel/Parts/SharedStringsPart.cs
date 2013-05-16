using System;
using System.Net;
using System.Windows;
using System.Windows.Input;
using FiftyNine.Ag.OpenXML.Common;
using FiftyNine.Ag.OpenXML.Common.Packaging;
using FiftyNine.Ag.OpenXML.Excel.Utilities;
using System.Collections.Generic;
using FiftyNine.Ag.OpenXML.Excel.Elements;
using System.Linq;

namespace FiftyNine.Ag.OpenXML.Excel.Parts
{
    public class SharedStringsPart : PackagePart
    {
        List<SharedStringDefinition> _strings = new List<SharedStringDefinition>();

        protected override void SaveContent(System.Xml.XmlWriter writer)
        {
            writer.WriteStartElement("sst", SpreadsheetNamespaces.Main.Value);
            int count = Strings.Count();
            writer.WriteAttributeString("count", count.ToString());
            writer.WriteAttributeString("uniqueCount", count.ToString());

            foreach (var str in Strings)
            {
                writer.WriteStartElement("si");
                writer.WriteElementString("t", str.String);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public SharedStringDefinition AddString(string str)
        {
            SharedStringDefinition def = _strings.FirstOrDefault(ss => ss.String == str);
            if (def == null)
            {
                def = new SharedStringDefinition() { String = str, Index = _strings.Count };
                _strings.Add(def);
            }
            return def;
        }

        public IEnumerable<SharedStringDefinition> Strings
        {
            get { return _strings; }
        }
        protected override string ContentType
        {
            get { return SpreadsheetContentTypes.SharedStrings; }
        }
        protected override string RelationshipType
        {
            get { return SpreadsheetRelationshipTypes.SharedStrings; }
        }
    }

    public class SharedStringDefinition
    {
        public int Index { get; internal set; }
        public string String { get; internal set; }
    }
}
