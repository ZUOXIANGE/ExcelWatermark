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

        var ordersTiled = Path.Combine(AppContext.BaseDirectory, "orders_tiled_" + Guid.NewGuid() + ".xlsx");
        WorkbookFactory.CreateSampleOrdersWorkbook(ordersTiled, 200);
        var bytesTiled = WatermarkImageGenerator.GenerateTiledWatermarkImage("TILED WM", 1200, 900, -30f, 0.18f, "Microsoft YaHei", 40f, 260, 180, "#0066CC");
        BackgroundWatermark.SetBackgroundImage(ordersTiled, "Orders", bytesTiled);
        Console.WriteLine("Orders tiled: " + ordersTiled);

        var ordersCentered = Path.Combine(AppContext.BaseDirectory, "orders_centered_" + Guid.NewGuid() + ".xlsx");
        WorkbookFactory.CreateSampleOrdersWorkbook(ordersCentered, 200);
        var bytesCentered = WatermarkImageGenerator.GenerateCenteredWatermarkImage("CENTER WM", 1200, 900, -20f, 0.22f, "Microsoft YaHei", 96f, "#333333");
        BackgroundWatermark.SetBackgroundImage(ordersCentered, "Orders", bytesCentered);
        Console.WriteLine("Orders centered: " + ordersCentered);

        var ordersShadow = Path.Combine(AppContext.BaseDirectory, "orders_shadow_" + Guid.NewGuid() + ".xlsx");
        WorkbookFactory.CreateSampleOrdersWorkbook(ordersShadow, 200);
        var bytesShadow = WatermarkImageGenerator.GenerateTiledWatermarkImageWithShadow("SHADOW WM", 1200, 900, -35f, 0.20f, "Microsoft YaHei", 42f, 280, 190, "#222222", "#000000", 2, 2);
        BackgroundWatermark.SetBackgroundImage(ordersShadow, "Orders", bytesShadow);
        Console.WriteLine("Orders shadow: " + ordersShadow);

        var ordersOverlay = Path.Combine(AppContext.BaseDirectory, "orders_overlay_" + Guid.NewGuid() + ".xlsx");
        WorkbookFactory.CreateSampleOrdersWorkbook(ordersOverlay, 200);
        var ovSrc = WatermarkImageGenerator.GenerateCenteredWatermarkImage("OVER", 400, 300, 0f, 0.6f, "Microsoft YaHei", 64f, "#FF0000");
        var bytesOverlay = WatermarkImageGenerator.GenerateOverlayWatermarkImage(ovSrc, 1200, 900, -15f, 0.35f, 0.8f, 10, -5);
        BackgroundWatermark.SetBackgroundImage(ordersOverlay, "Orders", bytesOverlay);
        Console.WriteLine("Orders overlay: " + ordersOverlay);
    }
}