# EasyEPPlus

EPPlus extension

## Packages

--------
| Package | NuGet |
| ------- | ------------ |
| [EasyEPPlus](https://www.nuget.org/packages/EasyEPPlus/) | [![EasyEPPlus](https://img.shields.io/nuget/v/EasyEPPlus.svg)](https://www.nuget.org/packages/EasyEPPlus/) |

## Method

### EPPlusHeaderAttribute

``` csharp

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

```

### WriteToExcelAsync

``` csharp

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

List<TestDto> testDtos = new List<TestDto>();

string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test.xlsx");

await testDtos.WriteToExcelAsync(path);

```

### AppendToExcelAsync

``` csharp

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

List<TestDto> testDtos = new List<TestDto>();

string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test.xlsx");

await testDtos.AppendToExcelAsync(path);

```

### ReadFromExcel

``` csharp

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test.xlsx");

EPPlusExtensions.ReadFromExcel<TestDto>(path);

```

## Test

### Dto

``` csharp

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

```

### Test

``` csharp

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

```

### xlsx

![image](https://github.com/xybii/EasyEPPlus/blob/main/test.png)