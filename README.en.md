# ExcelWatermark

[Chinese](README.md) | [English](README.en.md)

An Excel watermark library based on .NET 10 and OpenXML, providing two capabilities:
- Background watermark: generate a diagonally tiled, semi‑transparent text PNG and attach it to a specified worksheet as a background image reference
- Blind watermark: encode encrypted text into workbook styles (font‑name bit and color LSB), hidden in the worksheet and safely extractable

## Features
- Background watermark
  - Generate tiled text PNG: `WatermarkImageGenerator.GenerateTiledWatermarkImage`
  - Set PNG as background: `BackgroundWatermark.SetBackgroundImage` (supports `filePath` and `Stream`)
- Blind watermark
  - Embed: `BlindWatermark.EmbedBlindWatermark` (`ExcelWatermark/BlindWatermark.cs:127`) encodes AES‑GCM encrypted payload into hidden sheet styles
  - Extract: `BlindWatermark.ExtractBlindWatermark` (`ExcelWatermark/BlindWatermark.cs:180`) reads styles, reconstructs the bitstream and decrypts
- Sample helpers (not published to NuGet)
  - Create blank workbook: `WorkbookFactory.CreateBlankWorkbook` (`ExcelWatermark.Sample/WorkbookFactory.cs:9`)
  - Create sample orders workbook: `WorkbookFactory.CreateSampleOrdersWorkbook` (`ExcelWatermark.Sample/WorkbookFactory.cs:24`)

## Project Layout
- `ExcelWatermark`: main library containing `BackgroundWatermark` and `BlindWatermark`
- `ExcelWatermark.Sample`: sample project demonstrating both watermark types, contains `WorkbookFactory`
- `ExcelWatermark.Tests`: xUnit test suite covering regressions and error scenarios, includes test `WorkbookFactory`
- `ExcelWatermark.slnx`: solution file

## Installation
```bash
dotnet add package ExcelWatermark
```
Or add a package reference in your `csproj`. After publishing you can install directly from NuGet.

## Environment & Dependencies
- .NET SDK `10.0.100` or higher
- NuGet packages:
  - `DocumentFormat.OpenXml` (OpenXML operations)
  - `System.Drawing.Common` (generate PNG watermark image, Windows platform)
- Platform support: because `System.Drawing.Common` is used, generating the background watermark image is supported only on Windows 6.1+.

## Quick Start
1) Build the solution:
```bash
dotnet build
```
2) Run the sample:
```bash
cd ExcelWatermark/ExcelWatermark.Sample
dotnet run --
```
The sample prints two temporary file paths:
- A workbook with blind watermark embedded and extracted
- A workbook with a background text watermark applied to the Orders sheet

## Usage Examples
- Embed and extract blind watermark (see `ExcelWatermark.Sample/Program.cs`)
```csharp
var file = Path.Combine(Path.GetTempPath(), "sample_" + Guid.NewGuid() + ".xlsx");
WorkbookFactory.CreateBlankWorkbook(file);
BlindWatermark.EmbedBlindWatermark(file, "Sample blind watermark", "demo-key");
var text = BlindWatermark.ExtractBlindWatermark(file, "demo-key");
Console.WriteLine(text);
```
- Blind watermark (Stream)
```csharp
var file = Path.Combine(Path.GetTempPath(), "sample_stream_" + Guid.NewGuid() + ".xlsx");
WorkbookFactory.CreateBlankWorkbook(file);
using (var fs = new FileStream(file, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
{
    BlindWatermark.EmbedBlindWatermark(fs, "Sample blind watermark", "demo-key");
}
using (var fr = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
{
    var text = BlindWatermark.ExtractBlindWatermark(fr, "demo-key");
    Console.WriteLine(text);
}
```
- Background text watermark (generate PNG and set as background)
```csharp
var orders = Path.Combine(Path.GetTempPath(), "orders_" + Guid.NewGuid() + ".xlsx");
WorkbookFactory.CreateSampleOrdersWorkbook(orders, 500);
var bytes = WatermarkImageGenerator.GenerateTiledWatermarkImage(
    "CONFIDENTIAL",
    1600, 1200, -30f, 0.12f,
    "Microsoft YaHei", 48f,
    280, 180,
    "#FF0000");
BackgroundWatermark.SetBackgroundImage(orders, "Orders", bytes);
```
- Background text watermark (Stream)
```csharp
var orders = Path.Combine(Path.GetTempPath(), "orders_stream_" + Guid.NewGuid() + ".xlsx");
WorkbookFactory.CreateSampleOrdersWorkbook(orders, 200);
using (var fs = new FileStream(orders, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
{
    var bytes = WatermarkImageGenerator.GenerateTiledWatermarkImage("STREAM WM", 800, 600, -45f, 0.2f, "Microsoft YaHei", 32f, 200, 150, "#0000FF");
    BackgroundWatermark.SetBackgroundImage(fs, "Orders", bytes);
}
```
- Set background from a PNG file
```csharp
var png = Path.Combine(Path.GetTempPath(), "bg.png");
File.WriteAllBytes(png, WatermarkImageGenerator.GenerateTiledWatermarkImage("FILE-WM", 600, 400));
BackgroundWatermark.SetBackgroundImageFromFile(orders, "Orders", png);
```

## How it Works
- Blind watermark
  - Encryption & integrity: `AES-GCM` (passphrase derived to a 256‑bit key via `SHA256`), payload structure `nonce(12)` + `tag(16)` + `ciphertext`
  - Carrier channel: font‑name bit (`Calibri=0` / `Cambria=1`) and blue channel color LSB, 2 bits per cell; written to hidden sheet `wm$`
- Background watermark
  - GDI+ draws diagonally tiled text on a transparent PNG, controlling rotation, stride, font, color and opacity
  - Attach the PNG as an ImagePart to the worksheet and add a `Picture` reference in `Worksheet` to achieve the background effect

## Tests
Run all tests:
```bash
dotnet test
```
Coverage: blind watermark embed/extract regressions, hidden sheet existence, incorrect key failure, background watermark image generation, image reference presence (file and stream).

## Notes
- Background watermark image generation depends on Windows (`System.Drawing.Common`)
- Fonts must be available on the system; for Chinese scenarios `Microsoft YaHei` is recommended
- Changes to color LSB are generally invisible, but may be lost in certain export or style cleanup processes (for blind watermark)

## License
This project is for demonstration, learning and experimentation. For production use, align with your security requirements and implement stricter key management and tamper-resistance.

Repository: https://github.com/ZUOXIANGE/ExcelWatermark