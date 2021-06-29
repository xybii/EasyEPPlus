using EasyEPPlus;
using Newtonsoft.Json;
using NUnit.Framework;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace EasyEPPlusTest
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void WriteToExcelAsyncTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DateTime dateTime = new DateTime(2021, 1, 1);

            List<TestDto> testDtos = new List<TestDto>();

            for (int i = 1; i < 10001; i++)
            {
                testDtos.Add(new TestDto()
                {
                    Id = i,
                    Name = $"Name_{i}",
                    DisplayName = $"DisplayName_{i}",
                    Created = dateTime.AddDays(i)
                });
            }

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Write.xlsx");

            testDtos.WriteToExcelAsync(path).GetAwaiter().GetResult();

            var dtos = EPPlusExtensions.ReadFromExcel<TestDto>(path);

            for (int i = 1; i < 10001; i++)
            {
                dtos[i - 1].DisplayName = $"DisplayName_{i}";
            }

            var t = JsonConvert.SerializeObject(testDtos) == JsonConvert.SerializeObject(dtos);

            Assert.True(t);
        }

        [Test]
        public void WriteToExcelAsyncTest2()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DateTime dateTime = new DateTime(2021, 1, 1);

            List<TestDto> testDtos = new List<TestDto>();

            for (int i = 1; i < 10001; i++)
            {
                testDtos.Add(new TestDto()
                {
                    Id = i,
                    Name = $"Name_{i}",
                    DisplayName = $"DisplayName_{i}",
                    Created = dateTime.AddDays(i)
                });
            }

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Write.xlsx");

            testDtos.WriteToExcelAsync(path, "Write").GetAwaiter().GetResult();

            var dtos = EPPlusExtensions.ReadFromExcel<TestDto>(path, "Read");

            Assert.IsNull(dtos);

            dtos = EPPlusExtensions.ReadFromExcel<TestDto>(path, "Write");

            Assert.IsNotNull(dtos);

            int index = 1;

            dtos.ForEach(o =>
            {
                if (index == 10001)
                {
                    index = 1;
                }

                o.DisplayName = $"DisplayName_{index++}";
            });

            var t = JsonConvert.SerializeObject(testDtos) == JsonConvert.SerializeObject(dtos);

            Assert.True(t);
        }

        [Test]
        public void AppendToExcelAsyncTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DateTime dateTime = new DateTime(2021, 1, 1);

            List<TestDto> testDtos = new List<TestDto>();

            for (int i = 1; i < 10001; i++)
            {
                testDtos.Add(new TestDto()
                {
                    Id = i,
                    Name = $"Name_{i}",
                    DisplayName = $"DisplayName_{i}",
                    Created = dateTime.AddDays(i)
                });
            }

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Append.xlsx");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            testDtos.AppendToExcelAsync(path).GetAwaiter().GetResult();

            testDtos.AppendToExcelAsync(path).GetAwaiter().GetResult();

            var dtos = EPPlusExtensions.ReadFromExcel<TestDto>(path);

            int index = 1;

            dtos.ForEach(o =>
            {
                if(index == 10001)
                {
                    index = 1;
                }

                o.DisplayName = $"DisplayName_{index++}";
            });

            testDtos.AddRange(testDtos);

            var t = JsonConvert.SerializeObject(testDtos) == JsonConvert.SerializeObject(dtos);

            Assert.True(t);
        }
    }
}