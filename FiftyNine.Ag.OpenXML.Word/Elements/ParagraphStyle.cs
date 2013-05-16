using System;
using System.Net;
using System.Windows;
using System.Windows.Input;
using FiftyNine.Ag.OpenXML.Common.Packaging;

namespace FiftyNine.Ag.OpenXML.Word.Elements
{
    public class ParagraphStyle : Style
    {
        public ParagraphStyle(PackagePart part) : base(part)
        {
            ParagraphProperties = Part.CreateElement<ParagraphProperties>();
        }

        protected override void SaveProperties(System.Xml.XmlWriter writer)
        {
            ParagraphProperties.Save(writer);
        }

        public override string Type
        {
            get { return "paragraph"; }
        }
        public ParagraphProperties ParagraphProperties
        {
            get;
            private set;
        }
    }
}
