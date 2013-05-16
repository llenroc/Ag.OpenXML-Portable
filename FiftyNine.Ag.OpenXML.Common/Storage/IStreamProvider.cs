using System;
using System.Net;
using System.IO;

namespace FiftyNine.Ag.OpenXML.Common.Storage
{
    public interface IStreamProvider : IDisposable
    {
        Stream GetStream(string path);
        void Close();
    }
}
