using EasyEPPlus;
using Newtonsoft.Json;
using OfficeOpenXml.Style;
using System;
using System.Drawing;

namespace EasyEPPlusTest
{
    public class TestDto
    {
        public int Id { get; set; }

        [JsonIgnore]
        public Image Loge { get; set; }

        [EPPlusHeader(UnderLine = true, ColorRGB = "38,28,220",
            HorizontalAlignment = ExcelHorizontalAlignment.Center,
            VerticalAlignment = ExcelVerticalAlignment.Center)]
        public Uri Url { get; set; }

        [EPPlusHeader(DisplayName = "TestName", Size = 20, Hyperlink = "https://www.qq.com")]
        public string Name { get; set; }

        [JsonIgnore]
        [EPPlusHeader(IsIgnore = true)]
        public string DisplayName { get; set; }

        [EPPlusHeader(Format = "yyyy:MM:dd HH-mm-ss", Bold = true, BackgroundColorRGB = "Tan")]
        public DateTime Created { get; set; }
    }
}
