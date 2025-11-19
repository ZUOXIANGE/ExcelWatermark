using System.Runtime.Versioning;
namespace ExcelWatermark.Sample;

class Program
{
    [SupportedOSPlatform("windows")]
    static void Main(string[] args)
    {
        var file = Path.Combine(AppContext.BaseDirectory, "sample_" + Guid.NewGuid() + ".xlsx");
        var key = "demo-key";
        var text = "示例盲水印";
        WorkbookFactory.CreateBlankWorkbook(file);
        BlindWatermark.EmbedBlindWatermark(file, text, key);
        var extracted = BlindWatermark.ExtractBlindWatermark(file, key);
        Console.WriteLine("Workbook: " + file);
        Console.WriteLine("Extracted: " + extracted);

        var orders = Path.Combine(AppContext.BaseDirectory, "orders_" + Guid.NewGuid() + ".xlsx");
        WorkbookFactory.CreateSampleOrdersWorkbook(orders, 500);
        var wmBytes = WatermarkImageGenerator.GenerateTiledWatermarkImage("CONFIDENTIAL 机密", 1600, 1200, -30f, 0.12f, "Microsoft YaHei", 48f, 280, 180, "#FF0000");
        BackgroundWatermark.SetBackgroundImage(orders, "Orders", wmBytes);
        Console.WriteLine("Orders workbook: " + orders);
    }
}