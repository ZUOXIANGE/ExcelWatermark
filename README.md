# ExcelWatermark

[![CI](https://github.com/ZUOXIANGE/ExcelWatermark/actions/workflows/nuget-publish.yml/badge.svg)](https://github.com/ZUOXIANGE/ExcelWatermark/actions/workflows/nuget-publish.yml)
[![NuGet](https://img.shields.io/nuget/v/ExcelWatermark.svg)](https://www.nuget.org/packages/ExcelWatermark)
[![Downloads](https://img.shields.io/nuget/dt/ExcelWatermark.svg)](https://www.nuget.org/packages/ExcelWatermark)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![.NET 10](https://img.shields.io/badge/.NET-10.0-purple.svg)](https://dotnet.microsoft.com/)

[中文](README.md) | [English](README.en.md)

一个基于 .NET 10 与 OpenXML 的 Excel 水印工具库，包含两类能力：
- 背景水印：生成斜向平铺的半透明文字 PNG，并追加到指定工作表作为背景图片引用
- 盲水印：将加密文本编码到工作簿样式（字体名位与颜色 LSB），隐藏在工作表中并可安全提取

## 功能
- 背景水印
  - 生成平铺文字的 PNG：`BackgroundWatermark.GenerateTiledWatermarkImage`（`ExcelWatermark/BackgroundWatermark.cs:27`）
  - 追加 PNG 为背景：`BackgroundWatermark.SetBackgroundImage`（支持 `filePath` 与 `Stream`，`ExcelWatermark/BackgroundWatermark.cs:80`、`ExcelWatermark/BackgroundWatermark.cs:101`）
  - 便捷方法：文字生成并设置背景 `SetBackgroundImageWithText`（`ExcelWatermark/BackgroundWatermark.cs:137`、`ExcelWatermark/BackgroundWatermark.cs:155`）
- 盲水印
  - 嵌入：`BlindWatermark.EmbedBlindWatermark`（`ExcelWatermark/BlindWatermark.cs:127`）将 AES‑GCM 加密负载编码到隐藏表样式
  - 提取：`BlindWatermark.ExtractBlindWatermark`（`ExcelWatermark/BlindWatermark.cs:180`）读取样式还原比特流并解密
- 示例辅助（不随 NuGet 发布）
  - 创建空工作簿：`WorkbookFactory.CreateBlankWorkbook`（`ExcelWatermark.Sample/WorkbookFactory.cs:9`）
  - 创建示例订单工作簿：`WorkbookFactory.CreateSampleOrdersWorkbook`（`ExcelWatermark.Sample/WorkbookFactory.cs:24`）

## 目录结构
- `ExcelWatermark`：主库（类库），包含 `BackgroundWatermark`、`BlindWatermark`
- `ExcelWatermark.Sample`：示例项目，演示两种水印的用法，包含 `WorkbookFactory`
- `ExcelWatermark.Tests`：xUnit 测试用例，覆盖回归与异常场景，包含测试用 `WorkbookFactory`
- `ExcelWatermark.slnx`：解决方案文件

## 安装
```bash
dotnet add package ExcelWatermark
```
或在 `csproj` 中添加对包的引用。发布后可直接通过 NuGet 安装。

## 环境与依赖
- .NET SDK `10.0.100` 及以上
- NuGet 包：
  - `DocumentFormat.OpenXml`（OpenXML 操作）
  - `System.Drawing.Common`（生成 PNG 水印图，Windows 平台）
- 平台支持：由于使用 `System.Drawing.Common`，背景水印图生成仅在 Windows 6.1+ 上受支持。

## 快速开始
1) 构建解决方案：
```bash
dotnet build
```
2) 运行示例：
```bash
cd ExcelWatermark/ExcelWatermark.Sample
dotnet run --
```
示例将输出两个临时文件路径：
- 一个嵌入并提取盲水印的工作簿
- 一个应用了背景文字水印的订单工作簿

## 使用示例
- 盲水印嵌入与提取（参考 `ExcelWatermark.Sample/Program.cs`）
```csharp
var file = Path.Combine(Path.GetTempPath(), "sample_" + Guid.NewGuid() + ".xlsx");
WorkbookFactory.CreateBlankWorkbook(file);
BlindWatermark.EmbedBlindWatermark(file, "示例盲水印", "demo-key");
var text = BlindWatermark.ExtractBlindWatermark(file, "demo-key");
Console.WriteLine(text);
```
- 背景文字水印（生成 PNG 并设置为背景）
```csharp
var orders = Path.Combine(Path.GetTempPath(), "orders_" + Guid.NewGuid() + ".xlsx");
WorkbookFactory.CreateSampleOrdersWorkbook(orders, 500);
BackgroundWatermark.SetBackgroundImageWithText(
    orders,
    "Orders",
    "CONFIDENTIAL 机密",
    1600, 1200, -30f, 0.12f,
    "Microsoft YaHei", 48f,
    280, 180,
    "#FF0000");
```
- 从 PNG 文件设置背景
```csharp
var png = Path.Combine(Path.GetTempPath(), "bg.png");
File.WriteAllBytes(png, BackgroundWatermark.GenerateTiledWatermarkImage("FILE-WM", 600, 400));
BackgroundWatermark.SetBackgroundImageFromFile(orders, "Orders", png);
```

## 原理简介
- 盲水印
  - 加密与完整性：`AES-GCM`（口令经 `SHA256` 派生为 256 位密钥），负载结构为 `nonce(12)` + `tag(16)` + `ciphertext`
  - 承载通道：字体名位（`Calibri=0` / `Cambria=1`）与字体颜色蓝通道 LSB 合计每格 2 位；写入隐藏工作表 `wm$`
- 背景水印
  - GDI+ 在透明 PNG 上斜向平铺绘制文字，控制旋转角度、步进、字体、颜色与不透明度
  - 将 PNG 作为 ImagePart 附加到工作表，并在 `Worksheet` 追加 `Picture` 引用以实现背景效果

## 测试
运行所有测试：
```bash
dotnet test
```
用例覆盖：盲水印嵌入/提取回归、隐藏表存在性、错误口令失败、背景水印图片生成、图片引用存在性（文件与流两种方式）。

## 注意事项
- 背景水印图片生成依赖 Windows（`System.Drawing.Common`）
- 字体名称需在系统可用；中文场景推荐 `Microsoft YaHei`
- 颜色 LSB 的改变一般不可见，但在某些导出或样式清洗场景下可能丢失（针对盲水印）

## 许可
本项目示例性质，适用于学习与实验用途。若用于生产环境，请结合你的业务安全需求进行更严格的密钥管理与抗篡改设计。

仓库地址：https://github.com/ZUOXIANGE/ExcelWatermark