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