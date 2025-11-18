using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.Versioning;

namespace ExcelWatermark;

/// <summary>
/// 背景水印工具类：
/// - 生成斜着铺满文字的半透明 PNG 水印图片
/// - 将 PNG 图片以关系引用的方式追加到目标工作表，实现背景水印效果
/// - 支持从文件或内存字节设置背景图片
/// </summary>
public static class BackgroundWatermark
{
    /// <summary>
    /// 生成斜向平铺文字的水印 PNG 图片字节。
    /// </summary>
    /// <param name="text">水印文字，支持中文</param>
    /// <param name="width">图片宽度（像素）</param>
    /// <param name="height">图片高度（像素）</param>
    /// <param name="angleDegrees">文字旋转角度（度），如 -30</param>
    /// <param name="opacity">不透明度 0~1，建议 0.1~0.2</param>
    /// <param name="fontFamily">字体名称，如 Microsoft YaHei</param>
    /// <param name="fontSize">字体大小（像素）</param>
    /// <param name="xStep">水平方向步进（间距）</param>
    /// <param name="yStep">垂直方向步进（间距）</param>
    /// <param name="colorHex">文字颜色，十六进制，如 #FF0000</param>
    [SupportedOSPlatform("windows")]
    public static byte[] GenerateTiledWatermarkImage(
        string text,
        int width = 1600,
        int height = 1200,
        float angleDegrees = -30f,
        float opacity = 0.15f,
        string fontFamily = "Microsoft YaHei",
        float fontSize = 36f,
        int xStep = 300,
        int yStep = 200,
        string colorHex = "#000000")
    {
        // 创建透明背景的位图与绘图对象
        using var bmp = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var g = System.Drawing.Graphics.FromImage(bmp);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
        g.Clear(System.Drawing.Color.Transparent);
        // 配置字体与半透明画刷
        using var font = new System.Drawing.Font(fontFamily, fontSize, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
        System.Drawing.Color baseColor = ParseHexColor(colorHex);
        var alpha = (int)Math.Round(Math.Clamp(opacity, 0f, 1f) * 255);
        using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(alpha, baseColor.R, baseColor.G, baseColor.B));
        // 将坐标系移至中心并旋转，便于斜向平铺
        g.TranslateTransform(width / 2f, height / 2f);
        g.RotateTransform(angleDegrees);
        var startX = -width;
        var endX = width;
        var startY = -height;
        var endY = height;
        // 双重循环，按步进绘制平铺文字
        for (int x = startX; x <= endX; x += xStep)
        {
            for (int y = startY; y <= endY; y += yStep)
            {
                g.DrawString(text, font, brush, x, y);
            }
        }
        g.ResetTransform();
        // 保存为 PNG 并返回字节
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        return ms.ToArray();
    }

    /// <summary>
    /// 生成文字水印图片并设置为指定工作表的背景。
    /// </summary>
    

    /// <summary>
    /// 将 PNG 图片追加到工作表并作为背景引用。
    /// </summary>
    public static void SetBackgroundImage(string filePath, string sheetName, byte[] imageBytes)
    {
        // 以可写方式打开工作簿并定位目标工作表
        using var doc = SpreadsheetDocument.Open(filePath, true);
        var wbPart = doc.WorkbookPart!;
        var wsPart = GetSheetByName(wbPart, sheetName) ?? throw new InvalidOperationException("Sheet not found");
        // 添加图片部件并写入字节
        var imagePart = wsPart.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(imageBytes))
        {
            imagePart.FeedData(stream);
        }
        // 获取图片关系 ID 并在工作表中追加 Picture 引用（作为背景）
        var relId = wsPart.GetIdOfPart(imagePart);
        var worksheet = wsPart.Worksheet;
        var picture = new Picture { Id = relId };
        worksheet.Append(picture);
        wsPart.Worksheet.Save();
        wbPart.Workbook.Save();
    }

    public static void SetBackgroundImage(Stream workbookStream, string sheetName, byte[] imageBytes)
    {
        using var doc = SpreadsheetDocument.Open(workbookStream, true);
        var wbPart = doc.WorkbookPart!;
        var wsPart = GetSheetByName(wbPart, sheetName) ?? throw new InvalidOperationException("Sheet not found");
        var imagePart = wsPart.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(imageBytes))
        {
            imagePart.FeedData(stream);
        }
        var relId = wsPart.GetIdOfPart(imagePart);
        var worksheet = wsPart.Worksheet;
        var picture = new Picture { Id = relId };
        worksheet.Append(picture);
        wsPart.Worksheet.Save();
        wbPart.Workbook.Save();
    }

    /// <summary>
    /// 从文件加载图片并设置为背景水印。
    /// </summary>
    public static void SetBackgroundImageFromFile(string filePath, string sheetName, string imageFilePath)
    {
        var bytes = File.ReadAllBytes(imageFilePath);
        SetBackgroundImage(filePath, sheetName, bytes);
    }

    public static void SetBackgroundImageFromFile(Stream workbookStream, string sheetName, string imageFilePath)
    {
        var bytes = File.ReadAllBytes(imageFilePath);
        SetBackgroundImage(workbookStream, sheetName, bytes);
    }

    /// <summary>
    /// 生成文字水印图片并设置为指定工作表的背景。
    /// </summary>
    [SupportedOSPlatform("windows")]
    public static void SetBackgroundImageWithText(
        string filePath,
        string sheetName,
        string text,
        int width = 1600,
        int height = 1200,
        float angleDegrees = -30f,
        float opacity = 0.15f,
        string fontFamily = "Microsoft YaHei",
        float fontSize = 36f,
        int xStep = 300,
        int yStep = 200,
        string colorHex = "#000000")
    {
        var bytes = GenerateTiledWatermarkImage(text, width, height, angleDegrees, opacity, fontFamily, fontSize, xStep, yStep, colorHex);
        SetBackgroundImage(filePath, sheetName, bytes);
    }

    [SupportedOSPlatform("windows")]
    public static void SetBackgroundImageWithText(
        Stream workbookStream,
        string sheetName,
        string text,
        int width = 1600,
        int height = 1200,
        float angleDegrees = -30f,
        float opacity = 0.15f,
        string fontFamily = "Microsoft YaHei",
        float fontSize = 36f,
        int xStep = 300,
        int yStep = 200,
        string colorHex = "#000000")
    {
        var bytes = GenerateTiledWatermarkImage(text, width, height, angleDegrees, opacity, fontFamily, fontSize, xStep, yStep, colorHex);
        SetBackgroundImage(workbookStream, sheetName, bytes);
    }

    /// <summary>
    /// 根据名称获取工作表对应的 WorksheetPart；不存在返回 null。
    /// </summary>
    private static WorksheetPart? GetSheetByName(WorkbookPart wbPart, string name)
    {
        var sheet = wbPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        if (sheet == null) return null;
        return (WorksheetPart)wbPart.GetPartById(sheet.Id!);
    }

    /// <summary>
    /// 解析十六进制颜色字符串为 System.Drawing.Color。
    /// 支持 "#RRGGBB" 与 "#AARRGGBB" 格式。
    /// </summary>
    private static System.Drawing.Color ParseHexColor(string hex)
    {
        hex = hex.Trim();
        if (hex.StartsWith("#")) hex = hex[1..];
        if (hex.Length == 6)
        {
            var r = Convert.ToByte(hex.Substring(0, 2), 16);
            var g = Convert.ToByte(hex.Substring(2, 2), 16);
            var b = Convert.ToByte(hex.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(255, r, g, b);
        }
        if (hex.Length == 8)
        {
            var a = Convert.ToByte(hex.Substring(0, 2), 16);
            var r = Convert.ToByte(hex.Substring(2, 2), 16);
            var g = Convert.ToByte(hex.Substring(4, 2), 16);
            var b = Convert.ToByte(hex.Substring(6, 2), 16);
            return System.Drawing.Color.FromArgb(a, r, g, b);
        }
        return System.Drawing.Color.Black;
    }
}