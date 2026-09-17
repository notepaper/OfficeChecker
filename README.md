# OfficeChecker

> [中文版](README.zh-CN.md) | English

One-click diagnostics for desktop Microsoft Office on Windows (WinForms, `net8.0-windows`).

Click the **检测 Office** ("Detect Office") button to run a six-layer check on
Excel / Word / PowerPoint, with the full report shown in the text box.

## Features

| # | Layer | Description |
|---|-------|-------------|
| 1 | Registry evidence | ProgID / CurVer / major version / install channel (MSI / ClickToRun) / registry bitness |
| 2 | PE evidence | Reads the EXE's PE header `Machine` field for the true bitness (x86 / x64 / arm64); takes precedence over the registry |
| 3 | File evidence | EXE FileVersion / ProductVersion / company name / product name |
| 4 | Live COM probing | Actually calls `CreateInstance` and reads `Name` / `Version` — the probe result is authoritative |
| 5 | Assembly evidence | Whether Interop / NetOffice assemblies load in the current .NET runtime (tri-state: Loaded / NotFound / LoadFailed) |
| 6 | Consistency check | Registry bitness ↔ PE bitness match check, with a warning (PE wins) on mismatch |

Design notes:

- No dependency on `Microsoft.Office.Interop.*`, NetOffice, or `dynamic` — hand-written `IDispatch` late binding.
- x64 process + x86 Office is **not** an error: Office COM servers are standalone
  EXEs reached via out-of-process COM, so the live probe result is what counts.
- `ProbeAutomationSmokeTest` goes one step further and verifies that basic
  document objects (Workbook / Documents / Presentations) can be created.

## Requirements

- Windows 10/11 (COM automation is Windows-only)
- .NET 8 SDK or Visual Studio 2022 17.8+
- No NuGet packages needed (`Microsoft.Win32.Registry` is built into `net8.0-windows`)

## Quick start

```powershell
dotnet build OfficeChecker.csproj -c Debug
dotnet run --project OfficeChecker.csproj
```

Or open `OfficeChecker.sln` in Visual Studio, press F5, then click 「检测 Office」.

## Sample output

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

## Reusing it in your own project

`OfficeDetector.cs` is a self-contained single file (namespace `OfficeProbe`).
Copy it into any `net8.0-windows` project:

```csharp
// Full report (recommended)
OfficeDetectionReport report = OfficeDetector.DetectFull("Excel");
if (report.ComReady)
{
    // COM works — trust the live probe
}

// Basic info only (does not start an Office process)
OfficeAppInfo info = OfficeDetector.Detect("Word");

// Document-level smoke test (creates, then closes, an empty document)
ComProbeResult smoke = OfficeDetector.ProbeAutomationSmokeTest("Excel.Application");
```

Valid arguments are `"Excel"` / `"Word"` / `"PowerPoint"`; anything else throws `ArgumentException`.

## FAQ

- **Interop / NetOffice shows NotFound?**
  Normal — it just means the project doesn't reference those assemblies.
  Late binding doesn't need them; only add the NuGet package if you want
  early binding (IntelliSense, compile-time checking).

- **My app is x64 but Office is x86 — is that a problem?**
  Not for out-of-process COM calls (a successful probe means you're fine).
  Only in-process scenarios (VSTO add-ins, DLL injection) and the ~2GB memory
  limit of 32-bit Excel need attention.

- **Probe fails with `0x8001010A`?**
  Office is busy (a modal dialog is waiting for input). Dismiss it and retry.

- **Bulk file processing is slow?**
  Prefer ClosedXML / Open XML SDK for bulk read/write (no Office process,
  no dialogs, no leftover processes). Keep COM for features that genuinely
  need Office itself (printing, macros, complex layout).

- **Microsoft does not support Office automation on unattended servers** —
  use the Open XML approach for server-side scenarios.

## Project layout

```text
OfficeChecker.csproj      Project file (net8.0-windows, WinExe)
OfficeChecker.sln         Solution
Program.cs                Entry point (STAThread)
Form1.cs                  Button-click logic (calls DetectFull, shows report)
Form1.Designer.cs         Form layout (button + result text box)
OfficeDetector.cs         The detection library (reusable standalone)
```
