using System;
using System.IO;
using FiftyNine.Ag.OpenXML.Common.Packaging;
using FiftyNine.Ag.OpenXML.Common.Storage;
using FiftyNine.Ag.OpenXML.Word.Constants;

namespace FiftyNine.Ag.OpenXML.Word.Parts
{
    public class ImagePart : PackagePart, IDisposable
    {
        Stream _img;

        protected override void SavePart(IStreamProvider storage)
        {
            using (Stream str = storage.GetStream(this.Path))
            {
                var buffer = new byte[_img.Length];
                _img.Read(buffer, 0, (int)_img.Length);
                str.Write(buffer, 0, buffer.Length);
            }
        }

        protected override string ContentType
        {
            get
            {
                return ContentTypes.Jpg;
            }
        }
        protected override string RelationshipType
        {
            get { return RelationshipTypes.Image; }
        }

        public Stream Image
        {
            get { return _img; }
            set
            {
                if (value == null)
                    throw new ArgumentException("Image cannot be null", "value");
                _img = value;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                _img.Close();
            }
            _img = null;
        }
        ~ImagePart()
        {
            Dispose(false);
        }
    }
}
