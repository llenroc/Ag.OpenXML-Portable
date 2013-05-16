using System;
using System.Net;
using ICSharpCode.SharpZipLib.Zip;
using System.IO;

namespace FiftyNine.Ag.OpenXML.Common.Storage
{
    public class ZipStreamProvider : IStreamProvider
    {
        ZipOutputStream _stream = null;
        TempStream _currentStream = null;
        bool _disposed = false;

        public ZipStreamProvider(Stream stream)
        {
            _stream = new ZipOutputStream(stream);
            _stream.SetLevel(9);
            _stream.UseZip64 = UseZip64.Off;
        }

        public System.IO.Stream GetStream(string path)
        {
            VerifyDisposed();
            if (_stream == null)
            {
                throw new Exception("Cannot get streams after the storage has been closed");
            }
            _currentStream = new TempStream(path.TrimStart('/'));
            _currentStream.Closing += Stream_Closing;
            return _currentStream;
        }
        public void Close()
        {
            VerifyDisposed();
            if (_currentStream != null)
            {
                _currentStream.Dispose();
                _currentStream = null;
            }

            if (_stream != null)
            {
                _stream.Finish();
                _stream.Close();
                _stream = null;
            }
        }

        void VerifyDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException("ZipStreamProvider");
        }
        void Stream_Closing(object sender, EventArgs e)
        {
            _currentStream.Closing -= Stream_Closing;

            if (_currentStream.Length == 0)
            {
                return;
            }

            _currentStream.Position = 0;

            ZipEntry entry = new ZipEntry(_currentStream.Name.TrimStart('/'));
            entry.Size = _currentStream.Length;
            _stream.PutNextEntry(entry);
            _currentStream.WriteTo(_stream);
        }

        public void Dispose()
        {
            Dispose(true);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                Close();
                _disposed = true;
            }
            _stream = null;
            _currentStream = null;
        }
        ~ZipStreamProvider()
        {
            Dispose(false);
        }

        private class TempStream : MemoryStream
        {
            public TempStream(string name)
            {
                Name = name;
            }

            protected override void Dispose(bool disposing) {
                OnClosing();
                base.Dispose(disposing);
            }

            public string Name
            {
                get;
                private set;
            }

            protected virtual void OnClosing()
            {
                if (Closing != null)
                    Closing(this, new EventArgs());
            }
            public event EventHandler Closing;
        }
    }
}
