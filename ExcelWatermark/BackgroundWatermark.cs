using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelWatermark;

/// <summary>
/// 背景水印工具类
/// </summary>
public static class BackgroundWatermark
{
    /// <summary>
    /// 将 PNG 图片追加到工作表并作为背景引用
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="sheetName"></param>
    /// <param name="imageBytes"></param>
    /// <exception cref="InvalidOperationException"></exception>
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

    /// <summary>
    /// 生成文字水印图片并设置为指定工作表的背景
    /// </summary>
    /// <param name="workbookStream"></param>
    /// <param name="sheetName"></param>
    /// <param name="imageBytes"></param>
    /// <exception cref="InvalidOperationException"></exception>
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
    /// 从文件加载图片并设置为背景水印
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="sheetName"></param>
    /// <param name="imageFilePath"></param>
    public static void SetBackgroundImageFromFile(string filePath, string sheetName, string imageFilePath)
    {
        var bytes = File.ReadAllBytes(imageFilePath);
        SetBackgroundImage(filePath, sheetName, bytes);
    }

    /// <summary>
    /// 从文件加载图片并设置为背景水印
    /// </summary>
    /// <param name="workbookStream"></param>
    /// <param name="sheetName"></param>
    /// <param name="imageFilePath"></param>
    public static void SetBackgroundImageFromFile(Stream workbookStream, string sheetName, string imageFilePath)
    {
        var bytes = File.ReadAllBytes(imageFilePath);
        SetBackgroundImage(workbookStream, sheetName, bytes);
    }

    /// <summary>
    /// 根据名称获取工作表对应的 WorksheetPart；不存在返回 null
    /// </summary>
    /// <param name="wbPart"></param>
    /// <param name="name"></param>
    /// <returns></returns>
    private static WorksheetPart? GetSheetByName(WorkbookPart wbPart, string name)
    {
        var sheet = wbPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        if (sheet == null) return null;
        return (WorksheetPart)wbPart.GetPartById(sheet.Id!);
    }

}