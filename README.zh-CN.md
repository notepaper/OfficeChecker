# OfficeChecker

> [English](README.md) | 中文

Windows 桌面版 Microsoft Office 环境检测工具（WinForms，`net8.0-windows`）。

点击「检测 Office」按钮，一键完成 Excel / Word / PowerPoint 的六层检测，
并在文本框中输出完整诊断报告。

## 功能

| # | 检测层 | 说明 |
|---|--------|------|
| 1 | 注册表证据 | ProgID / CurVer / 主版本 / 安装渠道（MSI / ClickToRun）/ 注册表位数 |
| 2 | PE 证据 | 读取 EXE 的 PE 头 `Machine` 字段，得到真实机器位数（x86 / x64 / arm64），优先级高于注册表 |
| 3 | 文件证据 | EXE 的 FileVersion / ProductVersion / 公司名 / 产品名 |
| 4 | COM 实测 | 真正 `CreateInstance`，读取 `Name` / `Version`，以实测结果为准 |
| 5 | 程序集证据 | Interop / NetOffice 程序集能否被当前 .NET 运行时加载（三态：Loaded / NotFound / LoadFailed） |
| 6 | 一致性校验 | 注册表位数 ↔ PE 位数是否一致，不一致时告警并以 PE 为准 |

实现特点：

- 不依赖 `Microsoft.Office.Interop.*`、NetOffice、`dynamic`，手写 `IDispatch` 后期绑定。
- x64 进程 + x86 Office **不算错误**：Office COM Server 是独立 EXE，跨进程通信，最终以 COM 实测结果为准。
- 附带 `ProbeAutomationSmokeTest`：进一步验证能否创建 Workbook / Documents / Presentations 等基础文档对象。

## 环境要求

- Windows 10/11（COM 自动化仅支持 Windows）
- .NET 8 SDK 或 Visual Studio 2022 17.8+
- 无需 NuGet 包（`Microsoft.Win32.Registry` 在 `net8.0-windows` 中内置）

## 快速开始

```powershell
dotnet build OfficeChecker.csproj -c Debug
dotnet run --project OfficeChecker.csproj
```

或用 Visual Studio 打开 `OfficeChecker.sln`，F5 运行，点击「检测 Office」。

## 输出示例

```text
进程位数: x64

[Excel]
  ProgID   : Excel.Application
  CurVer   : Excel.Application.16
  路径     : C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE  存在: True
  版本     : 16.0   内部版本: 16.0.19127.20800
  文件版本 : 16.0.19127.20800   产品版本: 16.0.19127.20800
  渠道     : ClickToRun   最终位数: x86
  位数校验 : 注册表位数与 PE 位数一致: x86
  COM 实测 : OK COM 调用成功
  Name     : Microsoft Excel 16.0
  Interop  : NotFound Microsoft.Office.Interop.Excel
  NetOffice: NotFound NetOfficeFw.Excel
...
```

## 在其他项目中复用

`OfficeDetector.cs` 是自包含单文件（命名空间 `OfficeProbe`），直接复制到
`net8.0-windows` 项目即可使用：

```csharp
// 完整报告（推荐）
OfficeDetectionReport report = OfficeDetector.DetectFull("Excel");
if (report.ComReady)
{
    // COM 可用，以实测为准
}

// 仅基础信息（不启动 Office 进程）
OfficeAppInfo info = OfficeDetector.Detect("Word");

// 文档级冒烟测试（会创建再关闭一个空文档）
ComProbeResult smoke = OfficeDetector.ProbeAutomationSmokeTest("Excel.Application");
```

`"Excel"` / `"Word"` / `"PowerPoint"` 为合法参数，其他值抛 `ArgumentException`。

## 常见问题

- **Interop / NetOffice 显示 NotFound？**
  正常。表示项目未引用这些程序集。后期绑定不需要它们；
  只有需要早绑定（智能提示、编译期检查）时才引用对应 NuGet。

- **程序是 x64，Office 是 x86，有影响吗？**
  进程外 COM 调用不受影响（实测 OK 即可放心）。
  只有进程内场景（VSTO 插件、DLL 注入）和 32 位 Excel 约 2GB 内存上限需要注意。

- **COM 实测失败 `0x8001010A`？**
  Office 正忙（有模态弹窗等待人工输入），关掉弹窗后重试。

- **批量读写文件很慢？**
  大批量读写建议用 ClosedXML / Open XML SDK（不起 Office 进程、无弹窗、
  无残留进程），COM 只留给必须用 Office 原生功能的场景（如打印、宏、复杂排版）。

- **微软官方不支持在无人值守服务器上做 Office 自动化**，
  服务端场景请直接使用 Open XML 方案。

## 项目结构

```text
OfficeChecker.csproj      项目文件（net8.0-windows， WinExe）
OfficeChecker.sln         解决方案
Program.cs                程序入口（STAThread）
Form1.cs                  检测按钮点击逻辑（调用 DetectFull 并展示报告）
Form1.Designer.cs         窗体布局（按钮 + 结果文本框）
OfficeDetector.cs         检测库本体（可独立复用）
```
