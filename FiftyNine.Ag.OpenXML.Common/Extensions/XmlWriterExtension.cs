using System;
using System.Net;
using System.Xml;

namespace FiftyNine.Ag.OpenXML.Common.Extensions
{
    public static class XmlWriterExtension
    {
        public static void WritePrefixedStartElement(this XmlWriter writer, Namespace ns, string nodename)
        {
            writer.WriteStartElement(ns.Prefix, nodename, ns.Value);
        }
        public static void WritePrefixedAttributeString(this XmlWriter writer, Namespace ns, string attributeName, string value)
        {
            writer.WriteAttributeString(ns.Prefix, attributeName, ns.Value, value);
        }
    }
}
