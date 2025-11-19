using Xunit;
using DocumentFormat.OpenXml.Packaging;
using System.Runtime.Versioning;

namespace ExcelWatermark.Tests;

public class BackgroundWatermarkTests
{
    // 用例说明：
    // 验证生成的水印图片字节为 PNG 格式，且内容非空（最小尺寸限制）。
    [Fact]
    [SupportedOSPlatform("windows")]
    public void Generate_Image_Returns_Png_Bytes()
    {
        // 生成指定参数的斜向文字水印图片字节
        var bytes = WatermarkImageGenerator.GenerateTiledWatermarkImage("WM", 400, 300, -30f, 0.15f, "Microsoft YaHei", 24f, 160, 120, "#333333");
        // 基本校验：字节不为空且长度合理
        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 100);
        // PNG 头校验（魔数）：89 50 4E 47 0D 0A 1A 0A
        Assert.Equal(new byte[]{0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A}, bytes.Take(8).ToArray());
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void Generate_Image_Invalid_Size_Should_Throw()
    {
        Assert.Throws<ArgumentException>(() =>
        {
            WatermarkImageGenerator.GenerateTiledWatermarkImage("WM", 0, 300);
        });
    }

    // 用例说明：
    // 基于文件路径调用设置文字背景水印后，工作表中应存在 Picture 引用元素。
    [Fact]
    [SupportedOSPlatform("windows")]
    public void Set_Background_Image_With_Text_FilePath()
    {
        var temp = Path.Combine(Path.GetTempPath(), "bgwm_fp_" + Guid.NewGuid() + ".xlsx");
        try
        {
            // 创建示例订单工作簿并设置文字背景水印
            WorkbookFactory.CreateSampleOrdersWorkbook(temp, 10);
            var bytes = WatermarkImageGenerator.GenerateTiledWatermarkImage("TEST-WM", 800, 600, -40f, 0.18f, "Microsoft YaHei", 28f, 200, 150, "#FF8800");
            BackgroundWatermark.SetBackgroundImage(temp, "Orders", bytes);
            // 打开只读并检查目标工作表是否存在图片引用
            using var doc = SpreadsheetDocument.Open(temp, false);
            var wb = doc.WorkbookPart!;
            var sheet = wb.Workbook.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault(s => s.Name == "Orders");
            Assert.NotNull(sheet);
            var ws = (WorksheetPart)wb.GetPartById(sheet.Id!);
            var pics = ws.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Picture>().Count();
            Assert.True(pics >= 1);
        }
        finally
        {
            // 清理临时文件
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    // 用例说明：
    // 基于图片文件路径调用设置背景水印后，工作表中应存在 Picture 引用元素。
    [Fact]
    [SupportedOSPlatform("windows")]
    public void Set_Background_Image_From_File_FilePath()
    {
        var tempWb = Path.Combine(Path.GetTempPath(), "bgwm_file_" + Guid.NewGuid() + ".xlsx");
        var tempPng = Path.Combine(Path.GetTempPath(), "bgwm_img_" + Guid.NewGuid() + ".png");
        try
        {
            // 生成 PNG 文件并在示例工作簿中应用为背景水印
            File.WriteAllBytes(tempPng, WatermarkImageGenerator.GenerateTiledWatermarkImage("FILE-WM", 600, 400));
            WorkbookFactory.CreateSampleOrdersWorkbook(tempWb, 5);
            BackgroundWatermark.SetBackgroundImageFromFile(tempWb, "Orders", tempPng);
            // 打开只读并检查 Picture 元素数量
            using var doc = SpreadsheetDocument.Open(tempWb, false);
            var wb = doc.WorkbookPart!;
            var sheet = wb.Workbook.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault(s => s.Name == "Orders");
            Assert.NotNull(sheet);
            var ws = (WorksheetPart)wb.GetPartById(sheet.Id!);
            var pics = ws.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Picture>().Count();
            Assert.True(pics >= 1);
        }
        finally
        {
            // 清理临时文件
            if (File.Exists(tempWb)) File.Delete(tempWb);
            if (File.Exists(tempPng)) File.Delete(tempPng);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void Set_Background_Image_From_File_Stream()
    {
        var tempWb = Path.Combine(Path.GetTempPath(), "bgwm_file_stream_" + Guid.NewGuid() + ".xlsx");
        var tempPng = Path.Combine(Path.GetTempPath(), "bgwm_img_stream_" + Guid.NewGuid() + ".png");
        try
        {
            File.WriteAllBytes(tempPng, WatermarkImageGenerator.GenerateTiledWatermarkImage("FILE-WM", 600, 400));
            WorkbookFactory.CreateSampleOrdersWorkbook(tempWb, 5);
            using (var fs = new FileStream(tempWb, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                BackgroundWatermark.SetBackgroundImageFromFile(fs, "Orders", tempPng);
            }
            using var doc = SpreadsheetDocument.Open(tempWb, false);
            var wb = doc.WorkbookPart!;
            var sheet = wb.Workbook.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault(s => s.Name == "Orders");
            Assert.NotNull(sheet);
            var ws = (WorksheetPart)wb.GetPartById(sheet.Id!);
            var pics = ws.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Picture>().Count();
            Assert.True(pics >= 1);
        }
        finally
        {
            if (File.Exists(tempWb)) File.Delete(tempWb);
            if (File.Exists(tempPng)) File.Delete(tempPng);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void Set_Background_Image_Invalid_Sheet_Should_Throw()
    {
        var temp = Path.Combine(Path.GetTempPath(), "bgwm_invalid_sheet_" + Guid.NewGuid() + ".xlsx");
        try
        {
            WorkbookFactory.CreateSampleOrdersWorkbook(temp, 5);
            var bytes = WatermarkImageGenerator.GenerateTiledWatermarkImage("WM", 400, 300);
            Assert.Throws<InvalidOperationException>(() =>
            {
                BackgroundWatermark.SetBackgroundImage(temp, "NotExists", bytes);
            });
        }
        finally
        {
            if (File.Exists(temp)) File.Delete(temp);
        }
    }
}