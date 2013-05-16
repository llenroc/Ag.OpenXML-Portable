using System;
using System.Net;

namespace FiftyNine.Ag.OpenXML.Common.Constants
{
    public static class Namespaces
    {
        public static readonly Namespace ContentTypes = new Namespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types");
        public static readonly Namespace PackageRelationships = new Namespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");
        public static readonly Namespace Relationships = new Namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        public static readonly Namespace ExtProps = new Namespace("extProp", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        public static readonly Namespace CoreProps = new Namespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        public static readonly Namespace Vml = new Namespace("v", "urn:schemas-microsoft-com:vml");

        public static class Purl
        {
            public static readonly Namespace Dc = new Namespace("dc", "http://purl.org/dc/elements/1.1/");
            public static readonly Namespace DcTerms = new Namespace("dcterms", "http://purl.org/dc/terms/");
            public static readonly Namespace DcmiType = new Namespace("dcmitype", "http://purl.org/dc/dcmitype/");
        }

        public static class W3C
        {
            public static readonly Namespace Xsi = new Namespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
        }
    }
}
