using EasyEPPlus;
using System;

namespace EasyEPPlusTest
{
    public class TestDto
    {
        public int Id { get; set; }

        [EPPlusHeader(DisplayName = "TestName")]
        public string Name { get; set; }

        [EPPlusHeader(IsIgnore = true)]
        public string DisplayName { get; set; }

        [EPPlusHeader(Format = "yyyy-MM-dd HH:mm:ss")]
        public DateTime Created { get; set; }
    }
}
