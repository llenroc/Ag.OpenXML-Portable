using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using FiftyNine.Ag.OpenXML.Common;
using FiftyNine.Ag.OpenXML.Common.Constants;
using FiftyNine.Ag.OpenXML.Common.Extensions;
using FiftyNine.Ag.OpenXML.Common.Packaging;

namespace FiftyNine.Ag.OpenXML.Word.Elements
{
    public class Picture : RunContent
    {
        string _id;
        Relationship _rel;
        Size _size;

        protected override void Initialize()
        {
            _id = Guid.NewGuid().ToString("N");
            Part.AddRequiredNamespace(Namespaces.Vml);
            Part.AddRequiredNamespace(Namespaces.Relationships);
        }

        public override void Save(System.Xml.XmlWriter writer)
        {
            if (Image == null)
            {
                throw new Exception("Image cannot be empty");
            }
            if (Size == Size.Empty)
            {
                throw new Exception("Size cannot be empty");
            }
            base.Save(writer);
        }
        protected override void SaveContent(System.Xml.XmlWriter writer)
        {
            writer.WritePrefixedStartElement(Namespaces.Vml, "shape");
            writer.WriteAttributeString("style", string.Format("width:{0}pt;height:{1}pt", _size.Width, _size.Height));

            writer.WritePrefixedStartElement(Namespaces.Vml, "imageData");
            writer.WriteAttributeString(Namespaces.Relationships.Prefix, "id", Namespaces.Relationships.Value, _rel.ID);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        protected override string NodeName
        {
            get { return "pict"; }
        }
        protected override Namespace Namespace
        {
            get { return Word.Constants.Namespaces.Main; }
        }
        public Relationship Image
        {
            get { return _rel; }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("value","Image cannot be empty");
                }
                _rel = value;
            }
        }
        public Size Size
        {
            get { return _size; }
            set
            {
                if (Size == Size.Empty)
                {
                    throw new ArgumentNullException("value", "Size cannot be empty");
                }
                _size = value;
            }
        }
    }
}
