using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace EasyEPPlus
{
    public static class EPPlusExtensions
    {
        public static async Task WriteToExcelAsync<T>(this IEnumerable<T> list, string path, string sheetName = "Sheet")
        {
            if (list == null || string.IsNullOrEmpty(path) || string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentException();
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            FileInfo fileInfo = new FileInfo(path);

            var props = typeof(T).GetProperties();

            Dictionary<string, EPPlusHeaderAttribute> dic = new Dictionary<string, EPPlusHeaderAttribute>();

            Dictionary<int, double> widthDic = new Dictionary<int, double>();

            Dictionary<int, double> heightDic = new Dictionary<int, double>();

            using (var package = new ExcelPackage(fileInfo))
            {
                using (var worksheet = package.Workbook.Worksheets.Add(sheetName))
                {
                    int columnCount = 0;

                    int left = 0;

                    int dim = worksheet.Dimension?.Rows ?? 0;

                    dim = dim == 0 ? 0 : dim - 1;

                    for (int j = 0; j < props.Length; j++)
                    {
                        var prop = props[j];

                        var ePPlusHeader = Attribute.GetCustomAttribute(prop, typeof(EPPlusHeaderAttribute), false) as EPPlusHeaderAttribute;

                        ePPlusHeader = ePPlusHeader ?? new EPPlusHeaderAttribute() { DisplayName = prop.Name };

                        dic.Add(prop.Name, ePPlusHeader);

                        if (ePPlusHeader.IsIgnore)
                        {
                            left -= 1;

                            continue;
                        }

                        columnCount++;

                        if (dim == 0)
                        {
                            worksheet.Cells[1, j + 1 + left].Value = string.IsNullOrEmpty(ePPlusHeader.DisplayName) ? prop.Name : ePPlusHeader.DisplayName;
                        }
                    }

                    for (int i = 0; i < list.Count(); i++)
                    {
                        left = 0;

                        for (int j = 0; j < props.Length; j++)
                        {
                            var prop = props[j];

                            var ePPlusHeader = dic[prop.Name];

                            if (ePPlusHeader.IsIgnore)
                            {
                                left -= 1;

                                continue;
                            }

                            object obj = prop.GetValue(list.ElementAt(i), null);

                            if (obj == null)
                            {
                                continue;
                            }

                            int row = i + 2 + dim;

                            int col = j + 1 + left;

                            using (var cells = worksheet.Cells[row, col])
                            {
                                if (obj is Image img)
                                {
                                    string s = $"{row}_{col}";

                                    var t = CreatePicture(worksheet, s, img, row, col);

                                    if (widthDic.ContainsKey(col))
                                    {
                                        widthDic[col] = t.Item1 > widthDic[col] ?
                                            t.Item1 :
                                            widthDic[col];
                                    }
                                    else
                                    {
                                        widthDic.Add(col, t.Item1);
                                    }

                                    if (heightDic.ContainsKey(row))
                                    {
                                        heightDic[row] = t.Item2 > widthDic[row] ?
                                            t.Item2 :
                                            widthDic[row];
                                    }
                                    else
                                    {
                                        heightDic.Add(row, t.Item2);
                                    }
                                }
                                else
                                {
                                    if (obj.GetType() == typeof(Uri))
                                    {
                                        cells.Hyperlink = new Uri(obj.ToString(), UriKind.Absolute);
                                    }
                                    else if (obj.GetType() == typeof(DateTime) && !string.IsNullOrEmpty(ePPlusHeader.Format) && DateTime.TryParse(obj.ToString(), out var dateTime))
                                    {
                                        cells.Value = dateTime.ToString(ePPlusHeader.Format);
                                    }
                                    else
                                    {
                                        cells.Value = obj;

                                        if (!string.IsNullOrEmpty(ePPlusHeader.Hyperlink))
                                        {
                                            cells.Hyperlink = new Uri(ePPlusHeader.Hyperlink, UriKind.Absolute);
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(ePPlusHeader.ColorRGB))
                                    {
                                        if (ePPlusHeader.ColorRGB.Contains(","))
                                        {
                                            string[] colors = ePPlusHeader.ColorRGB.Split(',');

                                            if (colors.Length == 3)
                                            {
                                                cells.Style.Font.Color.SetColor(Color.FromArgb(int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2])));
                                            }
                                        }
                                        else
                                        {
                                            cells.Style.Font.Color.SetColor(Color.FromName(ePPlusHeader.ColorRGB));
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(ePPlusHeader.BackgroundColorRGB))
                                    {
                                        cells.Style.Fill.PatternType = ExcelFillStyle.Solid;

                                        if (ePPlusHeader.BackgroundColorRGB.Contains(","))
                                        {
                                            string[] colors = ePPlusHeader.BackgroundColorRGB.Split(',');

                                            if (colors.Length == 3)
                                            {
                                                cells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2])));
                                            }
                                        }
                                        else
                                        {
                                            cells.Style.Fill.BackgroundColor.SetColor(Color.FromName(ePPlusHeader.BackgroundColorRGB));
                                        }
                                    }

                                    cells.Style.Font.Name = string.IsNullOrEmpty(ePPlusHeader.FontName) ? cells.Style.Font.Name : ePPlusHeader.FontName;
                                    cells.Style.Font.Bold = ePPlusHeader.Bold;
                                    cells.Style.Font.Size = ePPlusHeader.Size;
                                    cells.Style.Font.UnderLine = ePPlusHeader.UnderLine;
                                    cells.Style.HorizontalAlignment = ePPlusHeader.HorizontalAlignment;
                                    cells.Style.VerticalAlignment = ePPlusHeader.VerticalAlignment;
                                }
                            }
                        }
                    }

                    for (int i = 1; i < columnCount + 1; i++)
                    {
                        worksheet.Column(i).AutoFit();
                    }

                    foreach (var item in widthDic)
                    {
                        worksheet.Column(item.Key).Width = item.Value;
                    }

                    package.DoAdjustDrawings = false;

                    foreach (var item in heightDic)
                    {
                        worksheet.Row(item.Key).Height = item.Value;
                    }

                    await package.SaveAsync();
                }
            }
        }

        public static async Task AppendToExcelAsync<T>(this IEnumerable<T> list, string path, string sheetName = "Sheet")
        {
            if (list == null || string.IsNullOrEmpty(path) || string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentException();
            }

            FileInfo fileInfo = new FileInfo(path);

            var props = typeof(T).GetProperties();

            Dictionary<string, EPPlusHeaderAttribute> dic = new Dictionary<string, EPPlusHeaderAttribute>();

            Dictionary<int, double> widthDic = new Dictionary<int, double>();

            Dictionary<int, double> heightDic = new Dictionary<int, double>();

            using (var package = new ExcelPackage(fileInfo))
            {
                using (var worksheet = package.Workbook.Worksheets[sheetName] ?? package.Workbook.Worksheets.Add(sheetName))
                {
                    int columnCount = 0;

                    int left = 0;

                    int dim = worksheet.Dimension?.Rows ?? 0;

                    dim = dim == 0 ? 0 : dim - 1;

                    for (int j = 0; j < props.Length; j++)
                    {
                        var prop = props[j];

                        var ePPlusHeader = Attribute.GetCustomAttribute(prop, typeof(EPPlusHeaderAttribute), false) as EPPlusHeaderAttribute;

                        ePPlusHeader = ePPlusHeader ?? new EPPlusHeaderAttribute() { DisplayName = prop.Name };

                        dic.Add(prop.Name, ePPlusHeader);

                        if (ePPlusHeader.IsIgnore)
                        {
                            left -= 1;

                            continue;
                        }

                        columnCount++;

                        if (dim == 0)
                        {
                            worksheet.Cells[1, j + 1 + left].Value = string.IsNullOrEmpty(ePPlusHeader.DisplayName) ? prop.Name : ePPlusHeader.DisplayName;
                        }
                    }

                    for (int i = 0; i < list.Count(); i++)
                    {
                        left = 0;

                        for (int j = 0; j < props.Length; j++)
                        {
                            var prop = props[j];

                            var ePPlusHeader = dic[prop.Name];

                            if (ePPlusHeader.IsIgnore)
                            {
                                left -= 1;

                                continue;
                            }

                            object obj = prop.GetValue(list.ElementAt(i), null);

                            if (obj == null)
                            {
                                continue;
                            }

                            int row = i + 2 + dim;

                            int col = j + 1 + left;

                            using (var cells = worksheet.Cells[row, col])
                            {
                                if (obj is Image img)
                                {
                                    string s = $"{row}_{col}";

                                    var t = CreatePicture(worksheet, s, img, row, col);

                                    if (widthDic.ContainsKey(col))
                                    {
                                        widthDic[col] = t.Item1 > widthDic[col] ?
                                            t.Item1 :
                                            widthDic[col];
                                    }
                                    else
                                    {
                                        widthDic.Add(col, t.Item1);
                                    }

                                    if (heightDic.ContainsKey(row))
                                    {
                                        heightDic[row] = t.Item2 > widthDic[row] ?
                                            t.Item2 :
                                            widthDic[row];
                                    }
                                    else
                                    {
                                        heightDic.Add(row, t.Item2);
                                    }
                                }
                                else
                                {
                                    if (obj.GetType() == typeof(Uri))
                                    {
                                        cells.Hyperlink = new Uri(obj.ToString(), UriKind.Absolute);
                                    }
                                    else if (obj.GetType() == typeof(DateTime) && !string.IsNullOrEmpty(ePPlusHeader.Format) && DateTime.TryParse(obj.ToString(), out var dateTime))
                                    {
                                        cells.Value = dateTime.ToString(ePPlusHeader.Format);
                                    }
                                    else
                                    {
                                        cells.Value = obj;

                                        if (!string.IsNullOrEmpty(ePPlusHeader.Hyperlink))
                                        {
                                            cells.Hyperlink = new Uri(ePPlusHeader.Hyperlink, UriKind.Absolute);
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(ePPlusHeader.ColorRGB))
                                    {
                                        if (ePPlusHeader.ColorRGB.Contains(","))
                                        {
                                            string[] colors = ePPlusHeader.ColorRGB.Split(',');

                                            if (colors.Length == 3)
                                            {
                                                cells.Style.Font.Color.SetColor(Color.FromArgb(int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2])));
                                            }
                                        }
                                        else
                                        {
                                            cells.Style.Font.Color.SetColor(Color.FromName(ePPlusHeader.ColorRGB));
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(ePPlusHeader.BackgroundColorRGB))
                                    {
                                        cells.Style.Fill.PatternType = ExcelFillStyle.Solid;

                                        if (ePPlusHeader.BackgroundColorRGB.Contains(","))
                                        {
                                            string[] colors = ePPlusHeader.BackgroundColorRGB.Split(',');

                                            if (colors.Length == 3)
                                            {
                                                cells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2])));
                                            }
                                        }
                                        else
                                        {
                                            cells.Style.Fill.BackgroundColor.SetColor(Color.FromName(ePPlusHeader.BackgroundColorRGB));
                                        }
                                    }

                                    cells.Style.Font.Name = string.IsNullOrEmpty(ePPlusHeader.FontName) ? cells.Style.Font.Name : ePPlusHeader.FontName;
                                    cells.Style.Font.Bold = ePPlusHeader.Bold;
                                    cells.Style.Font.Size = ePPlusHeader.Size;
                                    cells.Style.Font.UnderLine = ePPlusHeader.UnderLine;
                                    cells.Style.HorizontalAlignment = ePPlusHeader.HorizontalAlignment;
                                    cells.Style.VerticalAlignment = ePPlusHeader.VerticalAlignment;
                                }
                            }
                        }
                    }

                    for (int i = 1; i < columnCount + 1; i++)
                    {
                        worksheet.Column(i).AutoFit();
                    }

                    foreach (var item in widthDic)
                    {
                        worksheet.Column(item.Key).Width = item.Value;
                    }

                    package.DoAdjustDrawings = false;

                    foreach (var item in heightDic)
                    {
                        worksheet.Row(item.Key).Height = item.Value;
                    }

                    await package.SaveAsync();
                }
            }
        }

        public static List<TEntity> ReadFromExcel<TEntity>(string path, string sheetName = null) where TEntity : class, new()
        {
            List<TEntity> nRet = null;

            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException();
            }

            if (!File.Exists(path))
            {
                throw new FileNotFoundException();
            }

            FileInfo fileInfo = new FileInfo(path);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = null;

                if (string.IsNullOrEmpty(sheetName) && package.Workbook.Worksheets.Count > 0)
                {
                    worksheet = package.Workbook.Worksheets.First();
                }

                if (!string.IsNullOrEmpty(sheetName))
                {
                    worksheet = package.Workbook.Worksheets.FirstOrDefault(o => o.Name == sheetName);
                }

                if (worksheet == null)
                {
                    return nRet;
                }

                using (worksheet)
                {
                    if (worksheet.Dimension == null)
                    {
                        return nRet;
                    }

                    DataTable dt = new DataTable(worksheet.Name);

                    for (int rowNum = 1; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                    {
                        if (dt.Columns.Count == 0)
                        {
                            for (int columnNum = 1; columnNum <= worksheet.Dimension.End.Column; columnNum++)
                            {
                                dt.Columns.Add(worksheet.Cells[rowNum, columnNum].Value.ToString().Trim(), typeof(string));
                            }

                            continue;
                        }

                        DataRow dr = dt.NewRow();

                        for (int columnNum = 1; columnNum <= dt.Columns.Count; columnNum++)
                        {
                            dr[columnNum - 1] = worksheet.Cells[rowNum, columnNum].Value?.ToString()?.Trim();
                        }

                        dt.Rows.Add(dr);
                    }

                    if (dt.Rows.Count >= 1)
                    {
                        nRet = new List<TEntity>();

                        Type type = typeof(TEntity);

                        var props = typeof(TEntity).GetProperties();

                        var ePPlusHeaderAttributes = new Dictionary<PropertyInfo, EPPlusHeaderAttribute>();

                        foreach (PropertyInfo pi in props)
                        {
                            if (!pi.CanWrite)
                            {
                                continue;
                            }

                            var ePPlusHeader = Attribute.GetCustomAttribute(pi, typeof(EPPlusHeaderAttribute), false) as EPPlusHeaderAttribute;

                            ePPlusHeader = ePPlusHeader ?? new EPPlusHeaderAttribute() { DisplayName = pi.Name };

                            if (!ePPlusHeader.IsIgnore)
                            {
                                ePPlusHeaderAttributes.Add(pi, ePPlusHeader);
                            }
                        }

                        foreach (DataRow dr in dt.Rows)
                        {
                            TEntity t = new TEntity();

                            for (int i = 0; i < ePPlusHeaderAttributes.Count; i++)
                            {
                                var item = ePPlusHeaderAttributes.ElementAt(i);

                                var tempName = string.IsNullOrEmpty(item.Value.DisplayName) ? item.Key.Name : item.Value.DisplayName;

                                if (dt.Columns.Contains(tempName))
                                {
                                    if (item.Key.GetMethod.ReturnParameter.ParameterType == typeof(Image))
                                    {
                                        var r = dr.Table.Rows.IndexOf(dr) + 1;

                                        var ttt = worksheet.Drawings;

                                        var excelPicture = worksheet.Drawings
                                            .Where(o => o.From.Row == r && o.From.Column == i)
                                            .FirstOrDefault() as OfficeOpenXml.Drawing.ExcelPicture;

                                        if (excelPicture != null)
                                        {
                                            item.Key.SetValue(t, excelPicture.Image, null);
                                        }
                                    }
                                    else if (dr[tempName] != DBNull.Value)
                                    {
                                        if (item.Key.GetMethod.ReturnParameter.ParameterType == typeof(DateTime) &&
                                            !string.IsNullOrEmpty(item.Value.Format) &&
                                            DateTime.TryParseExact(dr[tempName].ToString(), item.Value.Format, null, DateTimeStyles.None, out DateTime dateTime))
                                        {
                                            item.Key.SetValue(t, dateTime, null);
                                        }
                                        else if (item.Key.GetMethod.ReturnParameter.ParameterType == typeof(Uri))
                                        {
                                            item.Key.SetValue(t, new Uri(dr[tempName].ToString()), null);
                                        }
                                        else
                                        {
                                            object value = TypeDescriptor.GetConverter(item.Key.GetMethod.ReturnParameter.ParameterType).ConvertFrom(dr[tempName]);

                                            item.Key.SetValue(t, value, null);
                                        }
                                    }
                                }
                            }

                            nRet.Add(t);
                        }
                    }
                }
            }

            return nRet;
        }

        private static double GetWidthInPixels(ExcelRange cell)
        {
            double columnWidth = cell.Worksheet.Column(cell.Start.Column).Width;
            Font font = new Font(cell.Style.Font.Name, cell.Style.Font.Size, FontStyle.Regular);
            double pxBaseline = Math.Round(MeasureString("1234567890", font) / 10);
            return columnWidth * pxBaseline;
        }

        private static double GetHeightInPixels(ExcelRange cell)
        {
            double rowHeight = cell.Worksheet.Row(cell.Start.Row).Height;
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiY = graphics.DpiY;
                return rowHeight * (1.0 / 72.0) * dpiY;
            }
        }

        private static float MeasureString(string s, Font font)
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                return g.MeasureString(s, font, int.MaxValue, StringFormat.GenericTypographic).Width;
            }
        }

        private static Tuple<double, double> CreatePicture(ExcelWorksheet worksheet, string name, Image image, int firstRow, int firstColumn)
        {
            double columnWidth = GetWidthInPixels(worksheet.Cells[firstRow, firstColumn]);

            double rowHeight = GetHeightInPixels(worksheet.Cells[firstRow, firstColumn]);

            var w = image.Width / columnWidth * worksheet.Column(firstColumn).Width;

            var h = image.Height / rowHeight * worksheet.Row(firstRow).Height;

            var pic = worksheet.Drawings.AddPicture(name, image);

            pic.SetPosition(firstRow - 1, 0, firstColumn - 1, 0);

            pic.SetSize(image.Width, image.Height);

            return new Tuple<double, double>(w, h);
        }
    }
}
