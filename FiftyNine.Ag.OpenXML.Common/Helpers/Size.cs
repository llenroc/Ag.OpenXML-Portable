using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FiftyNine.Ag.OpenXML.Common.Helpers {
    public class Size {
        public Size(double width, double height) {
            Width = width;
            Height = height;
        }

        public double Width { get; set; }
        public double Height { get; set; }

        public static Size Empty {
            get {
                return new Size(double.NegativeInfinity, double.NegativeInfinity);
            }
        }
    }
}
