using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FiftyNine.Ag.OpenXML.Common.Helpers {
    public class Thickness {
        public Thickness(double uniformLength) {
            Bottom = Left = Right = Top = uniformLength;
        }

        public Thickness(double left, double top, double right, double bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }
        public double Bottom { get; set; }
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
    }
}
