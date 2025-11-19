using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace ExcelWatermark;

/// <summary>
/// 盲水印工具类
/// </summary>
public static class BlindWatermark
{
    /// <summary>
    /// 将加密后的盲水印嵌入到 Excel 文件。
    /// - 创建/获取样式集合，准备四种组合以承载 2 位（字体名位 + 颜色LSB 位）
    /// - 在隐藏工作表 <c>wm$</c> 按 32 列栅格写入若干空单元格并应用承载样式
    /// </summary>
    /// <param name="filePath">Excel 工作簿文件路径（可读写）。</param>
    /// <param name="text">要嵌入的盲水印文本（UTF-8 编码）。</param>
    /// <param name="key">加密口令，用于派生 AES-GCM 密钥。</param>
    public static void EmbedBlindWatermark(string filePath, string text, string key)
    {
        using var doc = SpreadsheetDocument.Open(filePath, true);
        EmbedCore(doc.WorkbookPart!, text, key);
    }

    /// <summary>
    /// 将加密后的盲水印嵌入到 Excel 工作簿流。
    /// 与 <see cref="EmbedBlindWatermark(string, string, string)"/> 等效，但以流作为输入。
    /// </summary>
    /// <param name="workbookStream">Excel 工作簿流（可读写，定位到开头）。</param>
    /// <param name="text">要嵌入的盲水印文本（UTF-8 编码）。</param>
    /// <param name="key">加密口令，用于派生 AES-GCM 密钥。</param>
    public static void EmbedBlindWatermark(Stream workbookStream, string text, string key)
    {
        using var doc = SpreadsheetDocument.Open(workbookStream, true);
        EmbedCore(doc.WorkbookPart!, text, key);
    }

    /// <summary>
    /// 从隐藏工作表 <c>wm$</c> 读取单元格样式，解码每格承载的两位并还原比特流，解析帧头后解密得到原文。
    /// 错误口令或数据不完整将导致解密失败或异常。
    /// </summary>
    /// <param name="filePath">Excel 工作簿文件路径（只读）。</param>
    /// <param name="key">加密口令，用于派生 AES-GCM 密钥。</param>
    /// <returns>提取得到的原始文本。</returns>
    /// <exception cref="InvalidOperationException">未找到水印工作表或水印数据不完整。</exception>
    /// <exception cref="System.Security.Cryptography.AuthenticationTagMismatchException">口令错误导致 GCM 标签校验失败。</exception>
    public static string ExtractBlindWatermark(string filePath, string key)
    {
        using var doc = SpreadsheetDocument.Open(filePath, false);
        return ExtractCore(doc.WorkbookPart!, key);
    }

    /// <summary>
    /// 从 Excel 工作簿流提取盲水印文本。
    /// 与 <see cref="ExtractBlindWatermark(string, string)"/> 等效，但以流作为输入。
    /// </summary>
    /// <param name="workbookStream">Excel 工作簿流（只读，定位到开头）。</param>
    /// <param name="key">加密口令，用于派生 AES-GCM 密钥。</param>
    /// <returns>提取得到的原始文本。</returns>
    /// <exception cref="InvalidOperationException">未找到水印工作表或水印数据不完整。</exception>
    /// <exception cref="System.Security.Cryptography.AuthenticationTagMismatchException">口令错误导致 GCM 标签校验失败。</exception>
    public static string ExtractBlindWatermark(Stream workbookStream, string key)
    {
        using var doc = SpreadsheetDocument.Open(workbookStream, false);
        return ExtractCore(doc.WorkbookPart!, key);
    }

    /// <summary>
    /// 核心嵌入逻辑：初始化样式与隐藏工作表，按位编码并写入栅格。
    /// </summary>
    /// <param name="wbPart">工作簿部件。</param>
    /// <param name="text">要嵌入的盲水印文本。</param>
    /// <param name="key">加密口令。</param>
    private static void EmbedCore(WorkbookPart wbPart, string text, string key)
    {
        var styles = wbPart.WorkbookStylesPart ?? wbPart.AddNewPart<WorkbookStylesPart>();
        // ReSharper disable once ConditionIsAlwaysTrueOrFalseAccordingToNullableAPIContract  去掉会报错
        if (styles.Stylesheet == null)
        {
            styles.Stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellFormats(new CellFormat()));
        }
        var cfg = EnsureStyleCombos(styles);
        var payload = Encrypt(text, key);
        var framed = Frame(payload);
        var bits = ToBits(framed).ToArray();
        var cellsNeeded = (bits.Length + 1) / 2;
        var wsPart = EnsureHiddenSheet(wbPart, "wm$");
        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;
        int cols = 32;
        int rows = (int)Math.Ceiling(cellsNeeded / (double)cols);
        int bitIdx = 0;
        for (int r = 1; r <= rows; r++)
        {
            var row = new Row { RowIndex = (uint)r };
            sheetData.Append(row);
            for (int c = 0; c < cols; c++)
            {
                if (bitIdx >= bits.Length) break;
                var a = bits[bitIdx++];
                var b = bitIdx < bits.Length ? bits[bitIdx++] : (byte)0;
                uint styleIndex = a switch
                {
                    0 => b == 0 ? cfg.cf00 : cfg.cf01,
                    _ => b == 0 ? cfg.cf10 : cfg.cf11
                };
                var cell = new Cell
                {
                    CellReference = ColumnName(c) + r,
                    DataType = CellValues.String,
                    CellValue = new CellValue(string.Empty),
                    StyleIndex = styleIndex
                };
                row.Append(cell);
            }
        }
        wsPart.Worksheet.Save();
        styles.Stylesheet.Save();
        wbPart.Workbook.Save();
    }

    /// <summary>
    /// 核心提取逻辑：读取隐藏工作表样式还原位流，并解密得到原文。
    /// </summary>
    /// <param name="wbPart">工作簿部件。</param>
    /// <param name="key">加密口令。</param>
    /// <returns>提取得到的原始文本。</returns>
    /// <exception cref="InvalidOperationException">未找到水印工作表或水印数据不完整。</exception>
    private static string ExtractCore(WorkbookPart wbPart, string key)
    {
        var styles = wbPart.WorkbookStylesPart!;
        var wsPart = GetSheetByName(wbPart, "wm$");
        if (wsPart == null) throw new InvalidOperationException("Watermark sheet not found");
        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;
        var bits = new List<byte>();
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var si = cell.StyleIndex?.Value ?? 0u;
                var cf = GetCellFormat(styles, si);
                var f = GetFont(styles, cf.FontId!.Value);
                var fname = f.Elements<FontName>().FirstOrDefault()?.Val?.Value ?? "";
                byte a = fname.Equals("Cambria", StringComparison.OrdinalIgnoreCase) ? (byte)1 : (byte)0;
                var rgb = f.Elements<Color>().FirstOrDefault()?.Rgb?.Value;
                byte b = 0;
                if (!string.IsNullOrEmpty(rgb) && rgb.Length == 8)
                {
                    var blue = Convert.ToByte(rgb.Substring(6, 2), 16);
                    b = (byte)(blue & 0x01);
                }
                bits.Add(a);
                bits.Add(b);
            }
        }
        var bytes = FromBits(bits);
        var hdr = ParseHeader(bytes);
        var total = hdr.headerBytes + hdr.length;
        if (bytes.Length < total) throw new InvalidOperationException("Incomplete watermark");
        var payload = new byte[hdr.length];
        Buffer.BlockCopy(bytes, hdr.headerBytes, payload, 0, hdr.length);
        return Decrypt(payload, key);
    }

    /// <summary>
    /// 使用 AES-GCM 加密文本，输出负载为 nonce(12) + tag(16) + ciphertext。
    /// key 由用户口令经 SHA256 派生，nonce 随机生成；tag 用于完整性校验。
    /// </summary>
    private static byte[] Encrypt(string text, string pass)
    {
        var key = SHA256.HashData(Encoding.UTF8.GetBytes(pass)); // 口令派生为 256 位密钥
        var nonce = RandomNumberGenerator.GetBytes(12);          // 12 字节随机 nonce
        var pt = Encoding.UTF8.GetBytes(text);                   // 明文 UTF8 编码
        var ct = new byte[pt.Length];                            // 密文缓冲区
        var tag = new byte[16];                                  // GCM 认证标签 16 字节
        using var gcm = new AesGcm(key, 16);                     // 指定标签长度的 AES-GCM
        gcm.Encrypt(nonce, pt, ct, tag);                         // 执行加密
        var res = new byte[nonce.Length + tag.Length + ct.Length]; // 拼接输出负载
        Buffer.BlockCopy(nonce, 0, res, 0, nonce.Length);
        Buffer.BlockCopy(tag, 0, res, nonce.Length, tag.Length);
        Buffer.BlockCopy(ct, 0, res, nonce.Length + tag.Length, ct.Length);
        return res;
    }

    /// <summary>
    /// 从负载中拆分 nonce/tag/ciphertext 并使用同一口令解密，返回原始文本。
    /// 若口令错误或数据被篡改，将抛出 AuthenticationTagMismatchException。
    /// </summary>
    private static string Decrypt(byte[] payload, string pass)
    {
        var key = SHA256.HashData(Encoding.UTF8.GetBytes(pass));
        var nonce = new byte[12];
        var tag = new byte[16];
        Buffer.BlockCopy(payload, 0, nonce, 0, 12);
        Buffer.BlockCopy(payload, 12, tag, 0, 16);
        var ct = new byte[payload.Length - 28];
        Buffer.BlockCopy(payload, 28, ct, 0, ct.Length);
        var pt = new byte[ct.Length];
        using var gcm = new AesGcm(key, 16);
        gcm.Decrypt(nonce, ct, tag, pt);
        return Encoding.UTF8.GetString(pt);
    }

    /// <summary>
    /// 将字节数组转换为 MSB 优先的比特流（每字节 8 位，高位在前）。
    /// </summary>
    private static IEnumerable<byte> ToBits(byte[] data)
    {
        foreach (var b in data)
        {
            for (int i = 7; i >= 0; i--)
                yield return (byte)((b >> i) & 1);
        }
    }

    /// <summary>
    /// 将比特流还原为字节数组（不足 8 位的末尾按 0 填充）。
    /// </summary>
    private static byte[] FromBits(IReadOnlyList<byte> bits)
    {
        var bytes = new byte[(bits.Count + 7) / 8];
        int bi = 0;
        for (int i = 0; i < bytes.Length; i++)
        {
            byte v = 0;
            for (int j = 0; j < 8; j++)
            {
                v <<= 1;
                if (bi < bits.Count) v |= bits[bi++];
            }
            bytes[i] = v;
        }
        return bytes;
    }

    /// <summary>
    /// 封装带帧头的字节序列："BMWM"(4) + 版本(1) + 长度(4, 小端) + 数据。
    /// 便于在提取时校验并截取完整负载。
    /// </summary>
    private static byte[] Frame(byte[] payload)
    {
        var magic = "BMWM"u8.ToArray();
        var ver = new byte[] { 1 };
        var len = BitConverter.GetBytes(payload.Length);
        if (BitConverter.IsLittleEndian == false) Array.Reverse(len);
        var res = new byte[magic.Length + ver.Length + len.Length + payload.Length];
        Buffer.BlockCopy(magic, 0, res, 0, magic.Length);
        Buffer.BlockCopy(ver, 0, res, magic.Length, ver.Length);
        Buffer.BlockCopy(len, 0, res, magic.Length + ver.Length, len.Length);
        Buffer.BlockCopy(payload, 0, res, magic.Length + ver.Length + len.Length, payload.Length);
        return res;
    }

    /// <summary>
    /// 解析帧头并返回数据长度与头部字节数；头部不含实际负载数据。
    /// </summary>
    private static (int length, int headerBytes) ParseHeader(byte[] bytes)
    {
        if (bytes.Length < 9) throw new InvalidOperationException("Invalid watermark");
        if (bytes[0] != (byte)'B' || bytes[1] != (byte)'M' || bytes[2] != (byte)'W' || bytes[3] != (byte)'M') throw new InvalidOperationException("Invalid watermark");
        int len = BitConverter.ToInt32(bytes, 5);
        if (BitConverter.IsLittleEndian == false)
        {
            var arr = bytes.Skip(5).Take(4).Reverse().ToArray();
            len = BitConverter.ToInt32(arr, 0);
        }
        return (len, 9);
    }

    /// <summary>
    /// 获取指定名称的工作表；若不存在则创建一个隐藏工作表作为水印载体。
    /// </summary>
    private static WorksheetPart EnsureHiddenSheet(WorkbookPart wbPart, string name)
    {
        var exists = GetSheetByName(wbPart, name);
        if (exists != null) return exists;
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new Worksheet(new SheetData());
        wsPart.Worksheet.Save();
        var sheets = wbPart.Workbook.Sheets ?? wbPart.Workbook.AppendChild(new Sheets());
        var sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId!.Value).DefaultIfEmpty(0u).Max() + 1;
        var relId = wbPart.GetIdOfPart(wsPart);
        var sheet = new Sheet { Name = name, SheetId = sheetId, Id = relId, State = SheetStateValues.Hidden };
        sheets.Append(sheet);
        wbPart.Workbook.Save();
        return wsPart;
    }

    /// <summary>
    /// 根据名称查找工作表并返回其 WorksheetPart；不存在返回 null。
    /// </summary>
    private static WorksheetPart? GetSheetByName(WorkbookPart wbPart, string name)
    {
        var sheet = wbPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        if (sheet == null) return null;
        return (WorksheetPart)wbPart.GetPartById(sheet.Id!);
    }

    /// <summary>
    /// 通过样式索引读取 CellFormat 对象。
    /// </summary>
    private static CellFormat GetCellFormat(WorkbookStylesPart styles, uint styleIndex)
    {
        var cfs = styles.Stylesheet.CellFormats!;
        return cfs.ElementAt((int)styleIndex) as CellFormat ?? cfs.Elements<CellFormat>().ElementAt((int)styleIndex);
    }

    /// <summary>
    /// 通过 FontId 读取 Font 对象。
    /// </summary>
    private static Font GetFont(WorkbookStylesPart styles, uint fontId)
    {
        var fonts = styles.Stylesheet.Fonts!;
        return fonts.ElementAt((int)fontId) as Font ?? fonts.Elements<Font>().ElementAt((int)fontId);
    }

    /// <summary>
    /// 准备四种样式组合以承载 2 位：
    /// cf00: Calibri + 蓝色LSB=0；cf01: Calibri + 蓝色LSB=1；
    /// cf10: Cambria + 蓝色LSB=0；cf11: Cambria + 蓝色LSB=1。
    /// 字体颜色以 ARGB 写入，例如 FF000000 与 FF000001 的区别仅在蓝色最低位。
    /// </summary>
    private static (uint cf00, uint cf01, uint cf10, uint cf11) EnsureStyleCombos(WorkbookStylesPart styles)
    {
        var ss = styles.Stylesheet;
        var fonts = ss.Fonts ?? ss.AppendChild(new Fonts());
        var cellFormats = ss.CellFormats ?? ss.AppendChild(new CellFormats());
        if (!fonts.Elements<Font>().Any()) fonts.Append(new Font());
        if (!cellFormats.Elements<CellFormat>().Any()) cellFormats.Append(new CellFormat());
        var idxCalBlack0 = AddFont(fonts, "Calibri", "FF000000");
        var idxCalBlack1 = AddFont(fonts, "Calibri", "FF000001");
        var idxCamBlack0 = AddFont(fonts, "Cambria", "FF000000");
        var idxCamBlack1 = AddFont(fonts, "Cambria", "FF000001");
        var cf00 = AddCellFormat(cellFormats, idxCalBlack0);
        var cf01 = AddCellFormat(cellFormats, idxCalBlack1);
        var cf10 = AddCellFormat(cellFormats, idxCamBlack0);
        var cf11 = AddCellFormat(cellFormats, idxCamBlack1);
        fonts.Count = (uint)fonts.Elements<Font>().Count();
        cellFormats.Count = (uint)cellFormats.Elements<CellFormat>().Count();
        return ((uint)cf00, (uint)cf01, (uint)cf10, (uint)cf11);
    }

    /// <summary>
    /// 向样式表添加字体（包含名称与 ARGB 颜色）。返回新增字体索引。
    /// </summary>
    private static int AddFont(Fonts fonts, string name, string rgb)
    {
        var f = new Font();
        f.Append(new FontName { Val = name });
        f.Append(new Color { Rgb = HexBinaryValue.FromString(rgb) });
        fonts.Append(f);
        return fonts.Elements<Font>().Count() - 1;
    }

    /// <summary>
    /// 添加使用指定 FontId 的 CellFormat，并启用 ApplyFont 标记。返回样式索引。
    /// </summary>
    private static int AddCellFormat(CellFormats cfs, int fontId)
    {
        var cf = new CellFormat { FontId = (uint)fontId, ApplyFont = true };
        cfs.Append(cf);
        return cfs.Elements<CellFormat>().Count() - 1;
    }

    /// <summary>
    /// 将从 0 开始的列索引转换为 Excel 列名（如 0->A，27->AB）。
    /// </summary>
    private static string ColumnName(int index)
    {
        var sb = new StringBuilder();
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