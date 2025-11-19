using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace ExcelWatermark;

public static class WorkbookFactory
{
    public static void CreateBlankWorkbook(string filePath)
    {
        using var doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new Worksheet(new SheetData());
        var sheets = wbPart.Workbook.AppendChild(new Sheets());
        var relId = wbPart.GetIdOfPart(wsPart);
        var sheet = new Sheet { Name = "Sheet1", Id = relId, SheetId = 1 };
        sheets.Append(sheet);
        wbPart.Workbook.Save();
        wsPart.Worksheet.Save();
    }

    public static void CreateSampleOrdersWorkbook(string filePath, int rowCount = 1000)
    {
        using var doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        wsPart.Worksheet = new Worksheet(sheetData);

        var header = new Row { RowIndex = 1 };
        string[] headers = new[] { "OrderId", "CustomerName", "Product", "Category", "Quantity", "UnitPrice", "Total", "OrderDate", "Status", "Country", "City" };
        for (int i = 0; i < headers.Length; i++)
        {
            header.Append(new Cell { CellReference = ColumnName(i) + 1, DataType = CellValues.String, CellValue = new CellValue(headers[i]) });
        }
        sheetData.Append(header);

        var rnd = new Random(12345);
        string[] customers = new[] { "张三", "李四", "王五", "赵六", "孙七", "周八", "吴九", "郑十", "王洋", "刘敏" };
        string[] products = new[] { "手机", "笔记本电脑", "耳机", "键盘", "鼠标", "显示器", "路由器", "移动硬盘", "打印机", "平板" };
        string[] categories = new[] { "电子产品", "办公设备", "网络设备" };
        string[] statuses = new[] { "待支付", "已支付", "已发货", "已完成", "已取消" };
        string[] countries = new[] { "中国", "美国", "德国", "法国", "日本", "英国" };
        string[] citiesCN = new[] { "北京", "上海", "广州", "深圳", "杭州", "南京", "成都", "重庆", "西安", "武汉" };
        string[] citiesUS = new[] { "New York", "San Francisco", "Los Angeles", "Seattle", "Austin" };
        string[] citiesEU = new[] { "Berlin", "Munich", "Paris", "Lyon", "London", "Manchester" };

        for (int r = 0; r < rowCount; r++)
        {
            int rowIndex = r + 2;
            var row = new Row { RowIndex = (uint)rowIndex };

            string orderId = $"ORD-{DateTime.UtcNow:yyyyMMdd}-{r + 1:000000}";
            string customer = customers[rnd.Next(customers.Length)];
            string product = products[rnd.Next(products.Length)];
            string category = categories[rnd.Next(categories.Length)];
            int quantity = rnd.Next(1, 11);
            decimal unitPrice = Math.Round((decimal)rnd.NextDouble() * 495m + 5m, 2);
            decimal total = Math.Round(unitPrice * quantity, 2);
            DateTime date = DateTime.Today.AddDays(-rnd.Next(0, 365));
            string status = statuses[rnd.Next(statuses.Length)];
            string country = countries[rnd.Next(countries.Length)];
            string city = country switch
            {
                "中国" => citiesCN[rnd.Next(citiesCN.Length)],
                "美国" => citiesUS[rnd.Next(citiesUS.Length)],
                "英国" => citiesEU[rnd.Next(citiesEU.Length)],
                "德国" => citiesEU[rnd.Next(citiesEU.Length)],
                "法国" => citiesEU[rnd.Next(citiesEU.Length)],
                "日本" => "东京",
                _ => "Unknown"
            };

            row.Append(new Cell { CellReference = ColumnName(0) + rowIndex, DataType = CellValues.String, CellValue = new CellValue(orderId) });
            row.Append(new Cell { CellReference = ColumnName(1) + rowIndex, DataType = CellValues.String, CellValue = new CellValue(customer) });
            row.Append(new Cell { CellReference = ColumnName(2) + rowIndex, DataType = CellValues.String, CellValue = new CellValue(product) });
            row.Append(new Cell { CellReference = ColumnName(3) + rowIndex, DataType = CellValues.String, CellValue = new CellValue(category) });
            row.Append(new Cell { CellReference = ColumnName(4) + rowIndex, DataType = CellValues.Number, CellValue = new CellValue(quantity.ToString()) });
            row.Append(new Cell { CellReference = ColumnName(5) + rowIndex, DataType = CellValues.Number, CellValue = new CellValue(unitPrice.ToString(System.Globalization.CultureInfo.InvariantCulture)) });
            row.Append(new Cell { CellReference = ColumnName(6) + rowIndex, DataType = CellValues.Number, CellValue = new CellValue(total.ToString(System.Globalization.CultureInfo.InvariantCulture)) });
            row.Append(new Cell { CellReference = ColumnName(7) + rowIndex, DataType = CellValues.String, CellValue = new CellValue(date.ToString("yyyy-MM-dd")) });
            row.Append(new Cell { CellReference = ColumnName(8) + rowIndex, DataType = CellValues.String, CellValue = new CellValue(status) });
            row.Append(new Cell { CellReference = ColumnName(9) + rowIndex, DataType = CellValues.String, CellValue = new CellValue(country) });
            row.Append(new Cell { CellReference = ColumnName(10) + rowIndex, DataType = CellValues.String, CellValue = new CellValue(city) });

            sheetData.Append(row);
        }

        var sheets2 = wbPart.Workbook.AppendChild(new Sheets());
        var relId2 = wbPart.GetIdOfPart(wsPart);
        var sheet2 = new Sheet { Name = "Orders", Id = relId2, SheetId = 1 };
        sheets2.Append(sheet2);
        wbPart.Workbook.Save();
        wsPart.Worksheet.Save();
    }

    private static string ColumnName(int index)
    {
        var sb = new System.Text.StringBuilder();
        index++;
        while (index > 0)
        {
            var mod = (index - 1) % 26;
            sb.Insert(0, (char)('A' + mod));
            index = (index - 1) / 26;
        }
        return sb.ToString();
    }
}