using System;
using System.Net;
using FiftyNine.Ag.OpenXML.Common.Storage;
using System.IO;
using System.Xml;

namespace FiftyNine.Ag.OpenXML.Common.Extensions
{
    public static class IStreamProviderExtension
    {
        public static void CreateXMLPartFile(this IStreamProvider streamProvider, string filename, Action<XmlWriter> callback)
        {
            using (Stream str = streamProvider.GetStream(filename))
            using (XmlWriter writer = XmlWriter.Create(str))
            {
                writer.WriteStartDocument(true);
                callback(writer);
            }
        }
    }
}
