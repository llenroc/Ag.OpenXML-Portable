using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Xml;
using System.IO;
using System.Linq;

namespace FiftyNine.Ag.OpenXML.Common.Packaging
{
    public class Relationship
    {
        internal Relationship(PackagePart part, int relId) : this(part.Path, part.RelationshipType, false, relId)
        {
        }
        internal Relationship(string target, string type, bool isExternal, int relId)
        {
            ID = "r" + relId;
            Target = target;
            RelationshipType = type;
            IsExternal = isExternal;
        }

        protected internal void Save(XmlWriter writer, PackagePart relativeTo)
        {
            writer.WriteStartElement("Relationship");
            writer.WriteAttributeString("Id", ID);
            writer.WriteAttributeString("Type", RelationshipType);
            string target = Target;
            if (relativeTo != null)
            {
                string partPath = relativeTo.Path.Replace(Path.GetFileName(relativeTo.Path), "");
                if (target.StartsWith(partPath))
                {
                    target = target.Substring(partPath.Length);
                }
            }
            writer.WriteAttributeString("Target", target.TrimStart('/'));
            if (IsExternal)
            {
                writer.WriteAttributeString("TargetMode", "External");
            }
            writer.WriteEndElement();
        }

        public string ID
        {
            get;
            private set;
        }
        public string RelationshipType
        {
            get;
            private set;
        }
        public string Target
        {
            get;
            private set;
        }
        public bool IsExternal
        {
            get;
            private set;
        }
    }
}
