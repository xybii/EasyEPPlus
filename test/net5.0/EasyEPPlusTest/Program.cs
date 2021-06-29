using EasyEPPlus;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;

namespace EasyEPPlusTest
{
    class Program
    {
        static void Main(string[] args)
        {
            WriteToExcelAsyncTest();

            //WriteToExcelAsyncTest2();

            //AppendToExcelAsyncTest();

            Console.ReadKey();
        }

        public static void WriteToExcelAsyncTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DateTime dateTime = new DateTime(2021, 1, 1);

            List<TestDto> testDtos = new List<TestDto>();

            Image pic1 = Image.FromFile("pic1.jpg");

            Image pic2 = Image.FromFile("pic2.jpg");

            for (int i = 1; i < 101; i++)
            {
                testDtos.Add(new TestDto()
                {
                    Id = i,
                    Loge = i % 2 != 0 ? pic1 : pic2,
                    Url = new Uri("https://www.google.com/"),
                    Name = $"xx_test_{i}",
                    DisplayName = $"DisplayName_{i}",
                    Created = dateTime.AddDays(i)
                });
            }

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test.xlsx");

            Stopwatch stopwatch = Stopwatch.StartNew();

            testDtos.WriteToExcelAsync(path).GetAwaiter().GetResult();

            Console.WriteLine($"100 WriteToExcelAsync, {stopwatch.ElapsedMilliseconds}");

            var dtos = EPPlusExtensions.ReadFromExcel<TestDto>(path);

            var t = JsonConvert.SerializeObject(testDtos) == JsonConvert.SerializeObject(dtos);

            if (!t)
            {
                throw new Exception();
            }

            testDtos.AppendToExcelAsync(path).GetAwaiter().GetResult();

            testDtos.AddRange(testDtos);

            dtos = EPPlusExtensions.ReadFromExcel<TestDto>(path);

            t = JsonConvert.SerializeObject(testDtos) == JsonConvert.SerializeObject(dtos);

            if (!t)
            {
                throw new Exception();
            }
        }

        public static void WriteToExcelAsyncTest2()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DateTime dateTime = new DateTime(2021, 1, 1);

            List<TestDto> testDtos = new List<TestDto>();

            for (int i = 1; i < 101; i++)
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

            dtos = EPPlusExtensions.ReadFromExcel<TestDto>(path, "Write");

            var t = JsonConvert.SerializeObject(testDtos) == JsonConvert.SerializeObject(dtos);

            if (!t)
            {
                throw new Exception();
            }
        }

        public static void AppendToExcelAsyncTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DateTime dateTime = new DateTime(2021, 1, 1);

            List<TestDto> testDtos = new List<TestDto>();

            for (int i = 1; i < 101; i++)
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

            testDtos.AddRange(testDtos);

            var t = JsonConvert.SerializeObject(testDtos) == JsonConvert.SerializeObject(dtos);

            if (!t)
            {
                throw new Exception();
            }
        }
    }
}
