using OfficeOpenXml.Style;
using System;

namespace EasyEPPlus
{
    public class EPPlusHeaderAttribute : Attribute
    {
        public string DisplayName { get; set; }

        public string FontName { get; set; }

        public string Format { get; set; }

        public string Hyperlink { get; set; }

        public float Size { get; set; } = 15;

        public string ColorRGB { get; set; }

        public string BackgroundColorRGB { get; set; }

        public bool Bold { get; set; }

        public bool IsIgnore { get; set; }

        public bool UnderLine { get; set; }

        public ExcelHorizontalAlignment HorizontalAlignment { get; set; } = ExcelHorizontalAlignment.Left;

        public ExcelVerticalAlignment VerticalAlignment { get; set; } = ExcelVerticalAlignment.Bottom;
    }
}
