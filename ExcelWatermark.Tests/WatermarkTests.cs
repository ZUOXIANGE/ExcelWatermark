using Xunit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.Versioning;

namespace ExcelWatermark.Tests;

public class WatermarkTests
{
    // 用例说明：
    // 验证端到端流程：在临时工作簿中嵌入盲水印，再用同一口令提取，结果应与原文本一致。
    [Fact]
    public void Embed_And_Extract_Roundtrip()
    {
        // 准备一个临时文件路径，避免污染仓库
        var temp = Path.Combine(Path.GetTempPath(), "wm_rt_" + Guid.NewGuid() + ".xlsx");
        try
        {
            // 1) 创建空工作簿
            WorkbookFactory.CreateBlankWorkbook(temp);
            var text = "盲水印测试123ABC";
            var key = "mysecret";
            // 2) 嵌入盲水印（加密+位编码到隐藏工作表样式）
            BlindWatermark.EmbedBlindWatermark(temp, text, key);
            // 3) 提取盲水印（还原比特流并解密）
            var got = BlindWatermark.ExtractBlindWatermark(temp, key);
            // 4) 断言提取值等于原文
            Assert.Equal(text, got);
        }
        finally
        {
            // 清理临时文件
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    [Fact]
    public void Stream_Embed_And_Extract_Roundtrip()
    {
        var temp = Path.Combine(Path.GetTempPath(), "wm_rt_stream_" + Guid.NewGuid() + ".xlsx");
        try
        {
            WorkbookFactory.CreateBlankWorkbook(temp);
            var text = "Stream盲水印ABC123";
            var key = "stream-secret";
            using (var fs = new FileStream(temp, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                BlindWatermark.EmbedBlindWatermark(fs, text, key);
            }
            using (var fr = new FileStream(temp, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var got = BlindWatermark.ExtractBlindWatermark(fr, key);
                Assert.Equal(text, got);
            }
        }
        finally
        {
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    // 用例说明：
    // 验证嵌入后会创建名为 "wm$" 的隐藏工作表，作为水印载体。
    [Fact]
    public void Hidden_Sheet_Exists_After_Embed()
    {
        // 临时工作簿路径
        var temp = Path.Combine(Path.GetTempPath(), "wm_sheet_" + Guid.NewGuid() + ".xlsx");
        try
        {
            // 创建空工作簿并嵌入任意文本水印
            WorkbookFactory.CreateBlankWorkbook(temp);
            BlindWatermark.EmbedBlindWatermark(temp, "hello", "k");
            // 打开只读，检查工作簿中的 Sheets 集合
            using var doc = SpreadsheetDocument.Open(temp, false);
            var sheets = doc.WorkbookPart!.Workbook.Sheets!;
            // 查找名为 "wm$" 的工作表并确认其隐藏状态
            var sheet = sheets.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name, "wm$", StringComparison.OrdinalIgnoreCase));
            Assert.NotNull(sheet);
            Assert.Equal(SheetStateValues.Hidden, sheet.State!.Value);
        }
        finally
        {
            // 清理临时文件
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    [Fact]
    public void Extract_Without_Watermark_Sheet_Should_Fail()
    {
        var temp = Path.Combine(Path.GetTempPath(), "wm_no_sheet_" + Guid.NewGuid() + ".xlsx");
        try
        {
            WorkbookFactory.CreateBlankWorkbook(temp);
            Assert.Throws<InvalidOperationException>(() =>
            {
                BlindWatermark.ExtractBlindWatermark(temp, "k");
            });
        }
        finally
        {
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    [Fact]
    public void Extract_With_Empty_Watermark_Sheet_Should_Fail()
    {
        var temp = Path.Combine(Path.GetTempPath(), "wm_empty_" + Guid.NewGuid() + ".xlsx");
        try
        {
            WorkbookFactory.CreateBlankWorkbook(temp);
            using (var doc = SpreadsheetDocument.Open(temp, true))
            {
                var wb = doc.WorkbookPart!;
                var wsPart = wb.AddNewPart<WorksheetPart>();
                wsPart.Worksheet = new Worksheet(new SheetData());
                wsPart.Worksheet.Save();
                var sheets = wb.Workbook.Sheets ?? wb.Workbook.AppendChild(new Sheets());
                var sheetId = (uint)(sheets.Elements<Sheet>().Select(s => s.SheetId!.Value).DefaultIfEmpty(0u).Max() + 1);
                var relId = wb.GetIdOfPart(wsPart);
                var sheet = new Sheet { Name = "wm$", SheetId = sheetId, Id = relId, State = SheetStateValues.Hidden };
                sheets.Append(sheet);
                wb.Workbook.Save();
            }
            Assert.Throws<InvalidOperationException>(() =>
            {
                BlindWatermark.ExtractBlindWatermark(temp, "k");
            });
        }
        finally
        {
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    // 用例说明：
    // 验证错误口令提取时会失败，抛出 AES-GCM 的认证标签不匹配异常。
    [Fact]
    public void Extract_With_Wrong_Key_Should_Fail()
    {
        // 临时工作簿路径
        var temp = Path.Combine(Path.GetTempPath(), "wm_wrong_" + Guid.NewGuid() + ".xlsx");
        try
        {
            // 使用正确口令嵌入水印
            WorkbookFactory.CreateBlankWorkbook(temp);
            BlindWatermark.EmbedBlindWatermark(temp, "hello", "correct");
            // 使用错误口令提取，预期抛出 AuthenticationTagMismatchException
            Assert.Throws<System.Security.Cryptography.AuthenticationTagMismatchException>(() =>
            {
                BlindWatermark.ExtractBlindWatermark(temp, "wrong");
            });
        }
        finally
        {
            // 清理临时文件
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    [Fact]
    public void Create_Sample_Orders_Workbook_Should_Have_Many_Rows()
    {
        var temp = Path.Combine(Path.GetTempPath(), "orders_" + Guid.NewGuid() + ".xlsx");
        try
        {
            WorkbookFactory.CreateSampleOrdersWorkbook(temp, 200);
            using var doc = SpreadsheetDocument.Open(temp, false);
            var wb = doc.WorkbookPart!;
            var sheet = wb.Workbook.Sheets!.Elements<Sheet>().FirstOrDefault(s => s.Name == "Orders");
            Assert.NotNull(sheet);
            var ws = (WorksheetPart)wb.GetPartById(sheet.Id!);
            var rows = ws.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Count();
            Assert.True(rows >= 201);
        }
        finally
        {
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public void Stream_Background_Watermark_Adds_Picture()
    {
        var temp = Path.Combine(Path.GetTempPath(), "stream_wm_" + Guid.NewGuid() + ".xlsx");
        try
        {
            WorkbookFactory.CreateSampleOrdersWorkbook(temp, 10);
            using (var fs = new FileStream(temp, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                var bytes = WatermarkImageGenerator.GenerateTiledWatermarkImage("STREAM WM", 800, 600, -45f, 0.2f, "Microsoft YaHei", 32f, 200, 150, "#0000FF");
                BackgroundWatermark.SetBackgroundImage(fs, "Orders", bytes);
            }
            using var doc = SpreadsheetDocument.Open(temp, false);
            var wb = doc.WorkbookPart!;
            var sheet = wb.Workbook.Sheets!.Elements<Sheet>().FirstOrDefault(s => s.Name == "Orders");
            Assert.NotNull(sheet);
            var ws = (WorksheetPart)wb.GetPartById(sheet.Id!);
            var pics = ws.Worksheet.Elements<Picture>().Count();
            Assert.True(pics >= 1);
        }
        finally
        {
            if (File.Exists(temp)) File.Delete(temp);
        }
    }
}