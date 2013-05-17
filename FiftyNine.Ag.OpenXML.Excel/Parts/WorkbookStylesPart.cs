using FiftyNine.Ag.OpenXML.Common.Packaging;
using FiftyNine.Ag.OpenXML.Excel.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FiftyNine.Ag.OpenXML.Excel.Parts {
    public class WorkbookStylesPart : PackagePart {

        protected override void SaveContent(System.Xml.XmlWriter writer) {
            writer.WriteStartElement("styleSheet", SpreadsheetNamespaces.Main.Value);
            writer.WriteEndElement();
        }

        protected override string ContentType {
            get { return SpreadsheetContentTypes.WorkbookStyles; }
        }

        protected override string RelationshipType {
            get { return SpreadsheetRelationshipTypes.WorkbookStyles; }
        }
    }
}
