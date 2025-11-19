using System.Runtime.Versioning;

namespace ExcelWatermark;

/// <summary>
/// 水印图片生成器：负责生成斜向平铺文字的半透明 PNG 图像字节。
/// 不涉及工作表图片引用，配合 <see cref="BackgroundWatermark"/> 使用。
/// </summary>
public static class WatermarkImageGenerator
{
    /// <summary>
    /// 生成斜向平铺文字的水印 PNG 图片字节。
    /// </summary>
    /// <param name="text">水印文字，支持中文。</param>
    /// <param name="width">图片宽度（像素）。</param>
    /// <param name="height">图片高度（像素）。</param>
    /// <param name="angleDegrees">文字旋转角度（度），例如 -30。</param>
    /// <param name="opacity">不透明度 0~1，建议 0.1~0.2。</param>
    /// <param name="fontFamily">字体名称，例如 Microsoft YaHei。</param>
    /// <param name="fontSize">字体大小（像素）。</param>
    /// <param name="xStep">水平方向步进（间距，像素）。</param>
    /// <param name="yStep">垂直方向步进（间距，像素）。</param>
    /// <param name="colorHex">文字颜色十六进制，支持 "#RRGGBB" 或 "#AARRGGBB"。</param>
    /// <returns>PNG 格式图片字节数组。</returns>
    /// <remarks>依赖 System.Drawing.Common，仅在 Windows 上受支持。</remarks>
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
        using var bmp = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var g = System.Drawing.Graphics.FromImage(bmp);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
        g.Clear(System.Drawing.Color.Transparent);
        using var font = new System.Drawing.Font(fontFamily, fontSize, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
        System.Drawing.Color baseColor = ParseHexColor(colorHex);
        var alpha = (int)Math.Round(Math.Clamp(opacity, 0f, 1f) * 255);
        using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(alpha, baseColor.R, baseColor.G, baseColor.B));
        var measured = g.MeasureString(text, font);
        var stepX = Math.Max(xStep, (int)Math.Ceiling(measured.Width) + 4);
        var stepY = Math.Max(yStep, (int)Math.Ceiling(measured.Height) + 4);
        g.TranslateTransform(width / 2f, height / 2f);
        g.RotateTransform(angleDegrees);
        var startX = -width;
        var endX = width;
        var startY = -height;
        var endY = height;
        for (int x = startX; x <= endX; x += stepX)
        {
            for (int y = startY; y <= endY; y += stepY)
            {
                g.DrawString(text, font, brush, x, y);
            }
        }
        g.ResetTransform();
        using var ms = new System.IO.MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        return ms.ToArray();
    }

    /// <summary>
    /// 解析十六进制颜色字符串为 <see cref="System.Drawing.Color"/>。
    /// </summary>
    /// <param name="hex">十六进制颜色字符串，支持 "#RRGGBB" 与 "#AARRGGBB"。</param>
    /// <returns>解析得到的颜色；非法输入时返回黑色。</returns>
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