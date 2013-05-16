using System;
using System.Net;

namespace FiftyNine.Ag.OpenXML.Common
{
    public class Namespace
    {
        public Namespace(string prefix, string value)
        {
            Prefix = prefix;
            Value = value;
        }

        public string Prefix
        {
            get;
            private set;
        }
        public string Value
        {
            get;
            private set;
        }

        public override bool Equals(object obj)
        {
            if (obj is Namespace)
            {
                Namespace ns = (Namespace)obj;
                return ns.Prefix == Prefix && ns.Value == Value;
            }
            return false;
        }
        public override int GetHashCode()
        {
            return string.Format("{0}:{1}",Prefix,Value).GetHashCode();
        }
    }
}
