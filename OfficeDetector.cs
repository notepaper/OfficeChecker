// OfficeDetector.cs
//
// 目标框架:
//   net8.0-windows
//
// 设计目标:
//   1) 注册表证据：Office 是否安装、ProgID、版本、安装渠道
//   2) PE 证据：EXE 的真实机器位数
//   3) 文件证据：EXE FileVersion / ProductVersion
//   4) COM 实测：真正 CreateInstance，并访问 Name / Version
//   5) 程序集证据：Interop / NetOffice 是否能够被当前 .NET 运行时加载
//   6) 一致性校验：Registry Bitness <-> PE Bitness <-> COM Version
//
// 重要说明:
//   x64 .NET 进程 + x86 Office 并不天然是错误。
//   Office COM Server 是独立 EXE，通过 COM 跨进程通信。
//   因此是否兼容以实际 COM CreateInstance 结果为准，而不是单纯比较位数。
//
// 当前实现不依赖:
//   - Microsoft.Office.Interop.Excel
//   - Microsoft.Office.Interop.Word
//   - Microsoft.Office.Interop.PowerPoint
//   - NetOffice
//   - dynamic
//
// 适用:
//   - Office 环境诊断
//   - Tekla / WPF / WinForms / Console
//   - net8.0-windows

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OfficeProbe
{
    /// <summary>
    /// Office 应用基础信息。
    /// </summary>
    public sealed class OfficeAppInfo
    {
        public string Kind { get; init; } = string.Empty;

        public string ExeName { get; init; } = string.Empty;

        public string ProgId { get; init; } = string.Empty;

        /// <summary>
        /// Excel.Application.16 / Word.Application.16 ...
        /// </summary>
        public string? CurVer { get; set; }

        /// <summary>
        /// 根据 ProgID 推导出的主版本，例如 16.0。
        /// </summary>
        public string? Version { get; set; }

        /// <summary>
        /// 最终确认/推断得到的 EXE 路径。
        /// </summary>
        public string? ExePath { get; set; }

        /// <summary>
        /// App Paths\Path 对应的安装目录。
        /// </summary>
        public string? InstallDirectory { get; set; }

        /// <summary>
        /// 注册表推断的位数。
        /// x86 / x64 / arm64
        /// </summary>
        public string? RegistryBitness { get; set; }

        /// <summary>
        /// PE Header 实际位数。
        /// </summary>
        public string? PeBitness { get; set; }

        /// <summary>
        /// 最终位数。
        /// 优先使用 PE，注册表作为辅助证据。
        /// </summary>
        public string? Bitness { get; set; }

        /// <summary>
        /// MSI / ClickToRun / Unknown。
        /// </summary>
        public string Channel { get; set; } = "Unknown";

        /// <summary>
        /// Office / ClickToRun 级别版本。
        /// </summary>
        public string? Build { get; set; }

        /// <summary>
        /// EXE 文件版本。
        /// </summary>
        public string? FileVersion { get; set; }

        /// <summary>
        /// EXE ProductVersion。
        /// </summary>
        public string? ProductVersion { get; set; }

        /// <summary>
        /// 注册表位数与 PE 位数是否一致。
        /// </summary>
        public bool? BitnessConsistent { get; set; }

        /// <summary>
        /// EXE 文件是否存在。
        /// </summary>
        public bool ExeExists =>
            !string.IsNullOrWhiteSpace(ExePath) &&
            File.Exists(ExePath);
    }

    /// <summary>
    /// COM 检测结果。
    /// </summary>
    public sealed class ComProbeResult
    {
        /// <summary>
        /// 是否成功。
        /// </summary>
        public bool Ok { get; set; }

        /// <summary>
        /// ProgID 是否已注册。
        /// </summary>
        public bool ProgIdRegistered { get; set; }

        /// <summary>
        /// Name 属性。
        /// </summary>
        public string? DisplayName { get; set; }

        /// <summary>
        /// Version 属性。
        /// </summary>
        public string? Version { get; set; }

        /// <summary>
        /// HRESULT。
        /// 0 表示没有失败 HRESULT。
        /// </summary>
        public uint HResult { get; set; }

        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 具体异常类型。
        /// </summary>
        public string? ExceptionType { get; set; }

        /// <summary>
        /// 原始异常消息。
        /// </summary>
        public string? ExceptionMessage { get; set; }
    }

    /// <summary>
    /// .NET 程序集加载状态。
    /// </summary>
    public enum AssemblyLoadStatus
    {
        /// <summary>
        /// 未发现/无法解析程序集。
        /// </summary>
        NotFound,

        /// <summary>
        /// 找到并成功加载。
        /// </summary>
        Loaded,

        /// <summary>
        /// 运行时找到相关程序集，但加载失败。
        /// </summary>
        LoadFailed
    }

    /// <summary>
    /// .NET 程序集探测结果。
    /// </summary>
    public sealed class AssemblyProbeResult
    {
        public string AssemblyName { get; init; } = string.Empty;

        public AssemblyLoadStatus Status { get; init; }

        /// <summary>
        /// 是否成功加载。
        /// </summary>
        public bool Loaded =>
            Status == AssemblyLoadStatus.Loaded;

        /// <summary>
        /// 已加载程序集的实际路径。
        /// </summary>
        public string? Location { get; init; }

        public string? FullName { get; init; }

        public string? ErrorType { get; init; }

        public string? ErrorMessage { get; init; }
    }

    /// <summary>
    /// 文件版本信息。
    /// </summary>
    public sealed class FileVersionProbeResult
    {
        public bool Exists { get; init; }

        public string? FileVersion { get; init; }

        public string? ProductVersion { get; init; }

        public string? CompanyName { get; init; }

        public string? ProductName { get; init; }

        public string? Error { get; init; }
    }

    /// <summary>
    /// 位数一致性结果。
    /// </summary>
    public sealed class BitnessProbeResult
    {
        public string? RegistryBitness { get; init; }

        public string? PeBitness { get; init; }

        public bool? Consistent { get; init; }

        public string Message { get; init; } = string.Empty;
    }

    /// <summary>
    /// Office 完整检测报告。
    /// </summary>
    public sealed class OfficeDetectionReport
    {
        public OfficeAppInfo App { get; init; } = new();

        public ComProbeResult Com { get; init; } = new();

        public FileVersionProbeResult File { get; init; } = new();

        public BitnessProbeResult Bitness { get; init; } = new();

        public AssemblyProbeResult? InteropAssembly { get; init; }

        public AssemblyProbeResult? NetOfficeAssembly { get; init; }

        /// <summary>
        /// Office COM 环境是否可用。
        /// </summary>
        public bool ComReady =>
            Com.Ok;

        /// <summary>
        /// 当前是否存在可加载的 Interop。
        /// </summary>
        public bool InteropReady =>
            InteropAssembly?.Loaded == true;

        /// <summary>
        /// 当前是否存在可加载的 NetOffice。
        /// </summary>
        public bool NetOfficeReady =>
            NetOfficeAssembly?.Loaded == true;
    }

    public static class OfficeDetector
    {
        private sealed record OfficeDefinition(
            string Kind,
            string ExeName,
            string ProgId,
            string InteropAssemblyName,
            string NetOfficeAssemblyName);

        private static readonly OfficeDefinition[] KnownApps =
        {
            new(
                "Excel",
                "EXCEL.EXE",
                "Excel.Application",
                "Microsoft.Office.Interop.Excel",
                "NetOfficeFw.Excel"),

            new(
                "Word",
                "WINWORD.EXE",
                "Word.Application",
                "Microsoft.Office.Interop.Word",
                "NetOfficeFw.Word"),

            new(
                "PowerPoint",
                "POWERPNT.EXE",
                "PowerPoint.Application",
                "Microsoft.Office.Interop.PowerPoint",
                "NetOfficeFw.PowerPoint")
        };

        /// <summary>
        /// 完整检测。
        /// </summary>
        public static OfficeDetectionReport DetectFull(string kind)
        {
            var definition = GetDefinition(kind);

            var app = Detect(kind);

            var file = ProbeFileVersion(app.ExePath);

            var peBitness =
                !string.IsNullOrWhiteSpace(app.ExePath) &&
                File.Exists(app.ExePath)
                    ? PeMachine(app.ExePath)
                    : null;

            app.PeBitness = peBitness;

            app.FileVersion = file.FileVersion;
            app.ProductVersion = file.ProductVersion;

            app.BitnessConsistent =
                CompareBitness(
                    app.RegistryBitness,
                    app.PeBitness);

            // 最终位数以 PE 证据为优先；
            // 若 PE 不可用，则使用注册表。
            app.Bitness =
                app.PeBitness ??
                app.RegistryBitness;

            var bitnessProbe = new BitnessProbeResult
            {
                RegistryBitness = app.RegistryBitness,
                PeBitness = app.PeBitness,
                Consistent = app.BitnessConsistent,
                Message = BuildBitnessMessage(
                    app.RegistryBitness,
                    app.PeBitness)
            };

            var com = ProbeCom(app.ProgId);

            var interop = ProbeAssembly(
                definition.InteropAssemblyName);

            var netOffice = ProbeAssembly(
                definition.NetOfficeAssemblyName);

            return new OfficeDetectionReport
            {
                App = app,
                Com = com,
                File = file,
                Bitness = bitnessProbe,
                InteropAssembly = interop,
                NetOfficeAssembly = netOffice
            };
        }

        /// <summary>
        /// 仅执行 Office 基础安装信息检测。
        /// </summary>
        public static OfficeAppInfo Detect(string kind)
        {
            var definition = GetDefinition(kind);

            var info = new OfficeAppInfo
            {
                Kind = definition.Kind,
                ExeName = definition.ExeName,
                ProgId = definition.ProgId
            };

            // ------------------------------------------------------------
            // 1. App Paths
            // ------------------------------------------------------------

            string appPathsKey =
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" +
                definition.ExeName;

            foreach (var view in GetRegistryViews())
            {
                var defaultValue = ReadValue(
                    RegistryHive.LocalMachine,
                    appPathsKey,
                    null,
                    view);

                var pathValue = ReadValue(
                    RegistryHive.LocalMachine,
                    appPathsKey,
                    "Path",
                    view);

                if (!string.IsNullOrWhiteSpace(defaultValue))
                {
                    var exePath = StripQuotes(defaultValue);

                    if (File.Exists(exePath))
                    {
                        info.ExePath = exePath;

                        info.InstallDirectory =
                            Path.GetDirectoryName(exePath);

                        break;
                    }

                    // 默认值存在，但是文件不存在。
                    // 仍保留路径作为候选证据。
                    info.ExePath = exePath;
                }

                // 注意:
                // App Paths\Path 一般是“目录”，不是 EXE 完整路径。
                if (!string.IsNullOrWhiteSpace(pathValue))
                {
                    var directory =
                        StripQuotes(pathValue);

                    info.InstallDirectory = directory;

                    var candidate =
                        Path.Combine(
                            directory,
                            definition.ExeName);

                    if (File.Exists(candidate))
                    {
                        info.ExePath = candidate;
                        break;
                    }

                    if (string.IsNullOrWhiteSpace(info.ExePath))
                    {
                        info.ExePath = candidate;
                    }
                }
            }

            // ------------------------------------------------------------
            // 2. ProgID / CurVer
            // ------------------------------------------------------------

            info.CurVer =
                ReadClassesRootValue(
                    definition.ProgId + @"\CurVer");

            if (!string.IsNullOrWhiteSpace(info.CurVer))
            {
                info.Version =
                    ParseOfficeMajorVersion(info.CurVer);
            }

            // ------------------------------------------------------------
            // 3. 渠道 + 注册表位数
            //
            // ClickToRun 必须优先于 MSI。
            // 防止机器存在 MSI 残留信息时误判当前 Office 渠道。
            // ------------------------------------------------------------

            info.RegistryBitness =
                ReadBitness(
                    info.Version,
                    out var channel);

            info.Channel =
                channel ?? "Unknown";

            // ------------------------------------------------------------
            // 4. 如果前面没有获取到 EXE，再尝试根据版本/系统路径寻找
            // ------------------------------------------------------------

            if (string.IsNullOrWhiteSpace(info.ExePath) ||
                !File.Exists(info.ExePath))
            {
                var fallback =
                    FindExecutableFallback(
                        definition,
                        info.Version);

                if (!string.IsNullOrWhiteSpace(fallback))
                {
                    info.ExePath = fallback;
                    info.InstallDirectory =
                        Path.GetDirectoryName(fallback);
                }
            }

            // ------------------------------------------------------------
            // 5. Build
            // ------------------------------------------------------------

            info.Build =
                ReadBuild(
                    info.Version,
                    info.Channel);

            // ------------------------------------------------------------
            // 6. PE 位数
            // ------------------------------------------------------------

            if (!string.IsNullOrWhiteSpace(info.ExePath) &&
                File.Exists(info.ExePath))
            {
                info.PeBitness =
                    PeMachine(info.ExePath);

                info.Bitness =
                    info.PeBitness ??
                    info.RegistryBitness;

                info.BitnessConsistent =
                    CompareBitness(
                        info.RegistryBitness,
                        info.PeBitness);
            }
            else
            {
                info.Bitness =
                    info.RegistryBitness;
            }

            // ------------------------------------------------------------
            // 7. EXE FileVersion
            // ------------------------------------------------------------

            var version =
                ProbeFileVersion(info.ExePath);

            info.FileVersion = version.FileVersion;
            info.ProductVersion = version.ProductVersion;

            return info;
        }

        /// <summary>
        /// 获取应用定义。
        /// </summary>
        private static OfficeDefinition GetDefinition(string kind)
        {
            if (string.IsNullOrWhiteSpace(kind))
                throw new ArgumentException(
                    "Office 类型不能为空。",
                    nameof(kind));

            foreach (var item in KnownApps)
            {
                if (string.Equals(
                        item.Kind,
                        kind,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return item;
                }
            }

            throw new ArgumentException(
                "不支持的应用类型: " + kind,
                nameof(kind));
        }

        /// <summary>
        /// 注册表视图。
        /// </summary>
        private static RegistryView[] GetRegistryViews()
        {
            return new[]
            {
                RegistryView.Registry64,
                RegistryView.Registry32
            };
        }

        /// <summary>
        /// 读取 Office ClassesRoot。
        /// </summary>
        private static string? ReadClassesRootValue(
            string subKey)
        {
            foreach (var view in GetRegistryViews())
            {
                var value =
                    ReadValue(
                        RegistryHive.ClassesRoot,
                        subKey,
                        null,
                        view);

                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return null;
        }

        /// <summary>
        /// 获取 Office 位数与渠道。
        ///
        /// 优先级:
        ///   1. ClickToRun
        ///   2. MSI
        /// </summary>
        private static string? ReadBitness(
            string? version,
            out string? channel)
        {
            channel = null;

            // ------------------------------------------------------------
            // ClickToRun 优先
            // ------------------------------------------------------------

            foreach (var view in GetRegistryViews())
            {
                var platform =
                    ReadValue(
                        RegistryHive.LocalMachine,
                        @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                        "Platform",
                        view);

                if (!string.IsNullOrWhiteSpace(platform))
                {
                    channel = "ClickToRun";

                    return NormalizeBitness(
                        platform);
                }
            }

            // ------------------------------------------------------------
            // MSI
            // ------------------------------------------------------------

            if (!string.IsNullOrWhiteSpace(version))
            {
                foreach (var view in GetRegistryViews())
                {
                    var bitness =
                        ReadValue(
                            RegistryHive.LocalMachine,
                            @"SOFTWARE\Microsoft\Office\" +
                            version +
                            @"\Outlook",
                            "Bitness",
                            view);

                    if (!string.IsNullOrWhiteSpace(bitness))
                    {
                        channel = "MSI";

                        return NormalizeBitness(
                            bitness);
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 读取 Office Build。
        ///
        /// ClickToRun:
        ///   VersionToReport
        ///
        /// MSI:
        ///   Common\ProductVersion\LastProduct
        /// </summary>
        private static string? ReadBuild(
            string? version,
            string channel)
        {
            if (string.Equals(
                    channel,
                    "ClickToRun",
                    StringComparison.OrdinalIgnoreCase))
            {
                foreach (var view in GetRegistryViews())
                {
                    var c2r =
                        ReadValue(
                            RegistryHive.LocalMachine,
                            @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                            "VersionToReport",
                            view);

                    if (!string.IsNullOrWhiteSpace(c2r))
                        return c2r;
                }

                return null;
            }

            if (string.Equals(
                    channel,
                    "MSI",
                    StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(version))
            {
                foreach (var view in GetRegistryViews())
                {
                    var msi =
                        ReadValue(
                            RegistryHive.LocalMachine,
                            @"SOFTWARE\Microsoft\Office\" +
                            version +
                            @"\Common\ProductVersion",
                            "LastProduct",
                            view);

                    if (!string.IsNullOrWhiteSpace(msi))
                        return msi;
                }
            }

            return null;
        }

        /// <summary>
        /// 从 CurVer 中提取 Office 主版本。
        ///
        /// Excel.Application.16
        ///       ↓
        /// 16.0
        /// </summary>
        private static string? ParseOfficeMajorVersion(
            string curVer)
        {
            if (string.IsNullOrWhiteSpace(curVer))
                return null;

            var dot =
                curVer.LastIndexOf('.');

            if (dot <= 0 ||
                dot >= curVer.Length - 1)
            {
                return null;
            }

            var major =
                curVer[(dot + 1)..].Trim();

            return int.TryParse(
                major,
                out _)
                ? major + ".0"
                : null;
        }

        /// <summary>
        /// 规范化 Registry 位数。
        /// </summary>
        private static string? NormalizeBitness(
            string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            var v =
                value.Trim()
                    .ToLowerInvariant();

            return v switch
            {
                "x86" => "x86",
                "32-bit" => "x86",
                "32bit" => "x86",

                "x64" => "x64",
                "64-bit" => "x64",
                "64bit" => "x64",

                "arm64" => "arm64",

                _ => value.Trim()
            };
        }

        /// <summary>
        /// PE Machine 检测。
        ///
        /// 0x014C  = x86
        /// 0x8664  = x64
        /// 0xAA64  = ARM64
        /// </summary>
        public static string? PeMachine(
            string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) ||
                    !File.Exists(path))
                {
                    return null;
                }

                using var fs =
                    File.Open(
                        path,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite);

                using var br =
                    new BinaryReader(fs);

                if (fs.Length < 0x40)
                    return null;

                // DOS header: e_lfanew
                fs.Position = 0x3C;

                int peOffset =
                    br.ReadInt32();

                if (peOffset <= 0 ||
                    peOffset + 6 > fs.Length)
                {
                    return null;
                }

                fs.Position =
                    peOffset;

                var peSignature =
                    br.ReadUInt32();

                // "PE\0\0"
                if (peSignature != 0x00004550)
                    return null;

                // IMAGE_FILE_HEADER.Machine
                var machine =
                    br.ReadUInt16();

                return machine switch
                {
                    0x8664 => "x64",
                    0x014C => "x86",
                    0xAA64 => "arm64",
                    0x01C4 => "arm",
                    0xA641 => "arm64ec",

                    _ =>
                        "0x" +
                        machine.ToString("X4")
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// EXE 文件版本信息。
        /// </summary>
        private static FileVersionProbeResult ProbeFileVersion(
            string? path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    return new FileVersionProbeResult
                    {
                        Exists = false,
                        Error = "EXE 路径为空。"
                    };
                }

                if (!File.Exists(path))
                {
                    return new FileVersionProbeResult
                    {
                        Exists = false,
                        Error = "EXE 文件不存在: " + path
                    };
                }

                var info =
                    FileVersionInfo.GetVersionInfo(path);

                return new FileVersionProbeResult
                {
                    Exists = true,
                    FileVersion = info.FileVersion,
                    ProductVersion = info.ProductVersion,
                    CompanyName = info.CompanyName,
                    ProductName = info.ProductName
                };
            }
            catch (Exception ex)
            {
                return new FileVersionProbeResult
                {
                    Exists = false,
                    Error =
                        ex.GetType().Name +
                        ": " +
                        ex.Message
                };
            }
        }

        /// <summary>
        /// 注册表位数与 PE 位数一致性。
        /// </summary>
        private static bool? CompareBitness(
            string? registryBitness,
            string? peBitness)
        {
            if (string.IsNullOrWhiteSpace(registryBitness) ||
                string.IsNullOrWhiteSpace(peBitness))
            {
                return null;
            }

            return string.Equals(
                registryBitness,
                peBitness,
                StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildBitnessMessage(
            string? registryBitness,
            string? peBitness)
        {
            if (string.IsNullOrWhiteSpace(registryBitness) &&
                string.IsNullOrWhiteSpace(peBitness))
            {
                return "未获得位数证据。";
            }

            if (string.IsNullOrWhiteSpace(registryBitness))
            {
                return
                    "仅获得 PE 位数: " +
                    peBitness;
            }

            if (string.IsNullOrWhiteSpace(peBitness))
            {
                return
                    "仅获得注册表位数: " +
                    registryBitness;
            }

            if (string.Equals(
                    registryBitness,
                    peBitness,
                    StringComparison.OrdinalIgnoreCase))
            {
                return
                    $"注册表位数与 PE 位数一致: {peBitness}";
            }

            return
                $"位数证据冲突: Registry={registryBitness}, PE={peBitness}";
        }

        // ----------------------------------------------------------------
        // COM
        // ----------------------------------------------------------------

        /// <summary>
        /// 真正检测:
        ///
        /// ProgID
        ///   ↓
        /// Type.GetTypeFromProgID
        ///   ↓
        /// Activator.CreateInstance
        ///   ↓
        /// Name
        ///   ↓
        /// Version
        ///
        /// 这里不依赖 Office Interop。
        /// </summary>
        public static ComProbeResult ProbeCom(
            string progId)
        {
            var result =
                new ComProbeResult();

            Type? type = null;

            try
            {
#pragma warning disable CA1416
                type =
                    Type.GetTypeFromProgID(
                        progId,
                        false);
#pragma warning restore CA1416
            }
            catch (COMException ex)
            {
                result.Ok = false;
                result.ProgIdRegistered = false;
                result.HResult = unchecked((uint)ex.HResult);
                result.ExceptionType =
                    ex.GetType().Name;
                result.ExceptionMessage =
                    ex.Message;

                result.Message =
                    "GetTypeFromProgID 失败: " +
                    ExplainHResult(
                        result.HResult);

                return result;
            }
            catch (Exception ex)
            {
                result.Ok = false;
                result.ProgIdRegistered = false;
                result.ExceptionType =
                    ex.GetType().Name;
                result.ExceptionMessage =
                    ex.Message;

                result.Message =
                    "GetTypeFromProgID 异常: " +
                    ex.Message;

                return result;
            }

            if (type == null)
            {
                result.Ok = false;
                result.ProgIdRegistered = false;
                result.HResult = 0x800401F3;

                result.Message =
                    "ProgID 未注册: " +
                    progId;

                return result;
            }

            result.ProgIdRegistered = true;

            object? app = null;

            try
            {
                app =
                    Activator.CreateInstance(type);

                if (app == null)
                {
                    result.Ok = false;
                    result.Message =
                        "Activator.CreateInstance 返回 null。";

                    return result;
                }

                result.DisplayName =
                    ComLateBinding.GetProperty(
                        app,
                        "Name") as string;

                result.Version =
                    ComLateBinding.GetProperty(
                        app,
                        "Version") as string;

                result.Ok = true;
                result.HResult = 0;
                result.Message =
                    "COM 调用成功";
            }
            catch (COMException ex)
            {
                result.Ok = false;
                result.HResult =
                    unchecked((uint)ex.HResult);

                result.ExceptionType =
                    ex.GetType().Name;

                result.ExceptionMessage =
                    ex.Message;

                result.Message =
                    "CreateInstance 失败: " +
                    ExplainHResult(
                        result.HResult);
            }
            catch (Exception ex)
            {
                result.Ok = false;

                result.ExceptionType =
                    ex.GetType().Name;

                result.ExceptionMessage =
                    ex.Message;

                result.Message =
                    "COM 调用异常: " +
                    ex.GetType().Name +
                    " " +
                    ex.Message;
            }
            finally
            {
                if (app != null)
                {
                    try
                    {
                        ComLateBinding.InvokeMethod(
                            app,
                            "Quit");
                    }
                    catch
                    {
                        // 诊断过程中 Quit 失败不覆盖原始检测结果。
                    }

                    try
                    {
                        if (Marshal.IsComObject(app))
                        {
                            Marshal.FinalReleaseComObject(
                                app);
                        }
                    }
                    catch
                    {
                        // 释放失败同样不覆盖原始检测结果。
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// COM Smoke Test。
        ///
        /// 注意:
        /// ProbeCom() 只能证明:
        ///   COM Server 可以启动
        ///
        /// Smoke Test 进一步证明:
        ///   可以创建基础 Office 文档对象。
        ///
        /// 当前只对 Excel / Word / PowerPoint 做最小化测试。
        /// </summary>
        public static ComProbeResult ProbeAutomationSmokeTest(
            string progId)
        {
            var result =
                ProbeCom(progId);

            if (!result.Ok)
                return result;

            object? app = null;
            object? child = null;

            try
            {
                var type =
                    Type.GetTypeFromProgID(
                        progId,
                        false);

                if (type == null)
                {
                    result.Ok = false;
                    result.Message =
                        "Smoke Test: ProgID 未注册。";
                    return result;
                }

                app =
                    Activator.CreateInstance(type);

                if (app == null)
                {
                    result.Ok = false;
                    result.Message =
                        "Smoke Test: Office 实例创建失败。";
                    return result;
                }

                if (progId.Equals(
                        "Excel.Application",
                        StringComparison.OrdinalIgnoreCase))
                {
                    var workbooks =
                        ComLateBinding.GetProperty(
                            app,
                            "Workbooks");

                    if (workbooks == null)
                        throw new COMException(
                            "无法获取 Excel.Workbooks。");

                    child =
                        ComLateBinding.InvokeMethod(
                            workbooks,
                            "Add");

                    if (child == null)
                        throw new COMException(
                            "无法创建 Excel Workbook。");

                    try
                    {
                        ComLateBinding.InvokeMethod(
                            child,
                            "Close",
                            false);
                    }
                    catch
                    {
                        // Close 失败不重复覆盖原始信息。
                    }

                    if (Marshal.IsComObject(workbooks))
                    {
                        Marshal.FinalReleaseComObject(
                            workbooks);
                    }
                }
                else if (progId.Equals(
                             "Word.Application",
                             StringComparison.OrdinalIgnoreCase))
                {
                    child =
                        ComLateBinding.InvokeMethod(
                            app,
                            "Documents");

                    if (child == null)
                        throw new COMException(
                            "无法获取 Word.Documents。");
                }
                else if (progId.Equals(
                             "PowerPoint.Application",
                             StringComparison.OrdinalIgnoreCase))
                {
                    child =
                        ComLateBinding.GetProperty(
                            app,
                            "Presentations");

                    if (child == null)
                        throw new COMException(
                            "无法获取 PowerPoint.Presentations。");
                }

                result.Ok = true;
                result.Message =
                    "COM Automation Smoke Test 成功。";
            }
            catch (COMException ex)
            {
                result.Ok = false;
                result.HResult =
                    unchecked((uint)ex.HResult);

                result.ExceptionType =
                    ex.GetType().Name;

                result.ExceptionMessage =
                    ex.Message;

                result.Message =
                    "Smoke Test 失败: " +
                    ExplainHResult(
                        result.HResult);
            }
            catch (Exception ex)
            {
                result.Ok = false;

                result.ExceptionType =
                    ex.GetType().Name;

                result.ExceptionMessage =
                    ex.Message;

                result.Message =
                    "Smoke Test 异常: " +
                    ex.GetType().Name +
                    " " +
                    ex.Message;
            }
            finally
            {
                if (child != null)
                {
                    try
                    {
                        if (Marshal.IsComObject(child))
                        {
                            Marshal.FinalReleaseComObject(
                                child);
                        }
                    }
                    catch
                    {
                    }
                }

                if (app != null)
                {
                    try
                    {
                        ComLateBinding.InvokeMethod(
                            app,
                            "Quit");
                    }
                    catch
                    {
                    }

                    try
                    {
                        if (Marshal.IsComObject(app))
                        {
                            Marshal.FinalReleaseComObject(
                                app);
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// HRESULT 解释。
        ///
        /// 注意:
        /// 这里只描述“常见/可能原因”，
        /// 不把 HRESULT 直接等同于唯一根因。
        /// </summary>
        public static string ExplainHResult(
            uint hr)
        {
            return hr switch
            {
                0x80040154 =>
                    "REGDB_E_CLASSNOTREG：COM 类未注册。可能是 Office 未安装、COM 注册损坏或组件注册异常。",

                0x800401F3 =>
                    "CO_E_CLASSSTRING：ProgID 无效或无法解析。请检查 ProgID 注册。",

                0x80080005 =>
                    "CO_E_SERVER_EXEC_FAILURE：COM 服务器启动失败。可能与 Office 启动环境、用户配置、激活状态或服务器注册有关。",

                0x80070005 =>
                    "E_ACCESSDENIED：访问被拒绝。可能与权限、策略或 Office/Dcom 安全配置有关。",

                0x800706BA =>
                    "RPC_S_SERVER_UNAVAILABLE：RPC 服务器不可用。可能是服务器进程退出、RPC 状态异常或进程被终止。",

                0x8001010A =>
                    "RPC_E_SERVERCALL_RETRYLATER：Office 当前正忙，可能存在模态对话框或其他占用。",

                0x800A03EC =>
                    "Excel COM 内部错误。具体原因需要结合实际操作、文件、参数及上下文进一步判断。",

                _ =>
                    "0x" +
                    hr.ToString("X8")
            };
        }

        // ----------------------------------------------------------------
        // Assembly
        // ----------------------------------------------------------------

        /// <summary>
        /// 检测程序集。
        ///
        /// 与旧版 bool 方法相比，这里明确区分:
        ///
        /// NotFound
        /// LoadFailed
        /// Loaded
        ///
        /// 因此:
        /// “False” 不再吞掉所有错误信息。
        /// </summary>
        public static AssemblyProbeResult ProbeAssembly(
            string assemblyName)
        {
            if (string.IsNullOrWhiteSpace(assemblyName))
            {
                return new AssemblyProbeResult
                {
                    AssemblyName = assemblyName ?? string.Empty,
                    Status = AssemblyLoadStatus.NotFound,
                    ErrorType = nameof(ArgumentException),
                    ErrorMessage = "程序集名称为空。"
                };
            }

            try
            {
                var assembly =
                    Assembly.Load(
                        new AssemblyName(
                            assemblyName));

                return new AssemblyProbeResult
                {
                    AssemblyName = assemblyName,
                    Status = AssemblyLoadStatus.Loaded,
                    Location = assembly.Location,
                    FullName = assembly.FullName
                };
            }
            catch (FileNotFoundException ex)
            {
                return new AssemblyProbeResult
                {
                    AssemblyName = assemblyName,
                    Status = AssemblyLoadStatus.NotFound,
                    ErrorType = ex.GetType().Name,
                    ErrorMessage = ex.Message
                };
            }
            catch (FileLoadException ex)
            {
                return new AssemblyProbeResult
                {
                    AssemblyName = assemblyName,
                    Status = AssemblyLoadStatus.LoadFailed,
                    ErrorType = ex.GetType().Name,
                    ErrorMessage = ex.Message
                };
            }
            catch (BadImageFormatException ex)
            {
                return new AssemblyProbeResult
                {
                    AssemblyName = assemblyName,
                    Status = AssemblyLoadStatus.LoadFailed,
                    ErrorType = ex.GetType().Name,
                    ErrorMessage =
                        "程序集格式无效或架构不匹配: " +
                        ex.Message
                };
            }
            catch (Exception ex)
            {
                return new AssemblyProbeResult
                {
                    AssemblyName = assemblyName,
                    Status = AssemblyLoadStatus.LoadFailed,
                    ErrorType = ex.GetType().Name,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// 为兼容旧调用保留。
        ///
        /// 新代码建议使用 ProbeAssembly()。
        /// </summary>
        public static bool IsInteropLoaded(
            string assemblyName)
        {
            return ProbeAssembly(
                       assemblyName)
                   .Loaded;
        }

        // ----------------------------------------------------------------
        // 注册表 / 文件路径
        // ----------------------------------------------------------------

        private static string? ReadValue(
            RegistryHive hive,
            string subKey,
            string? name,
            RegistryView view)
        {
            try
            {
                using var baseKey =
                    RegistryKey.OpenBaseKey(
                        hive,
                        view);

                using var key =
                    baseKey.OpenSubKey(
                        subKey);

                if (key == null)
                    return null;

                var value =
                    key.GetValue(
                        name);

                if (value is not string str)
                    return null;

                return string.IsNullOrWhiteSpace(str)
                    ? null
                    : str.Trim();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 删除注册表 EXE 路径外层引号。
        /// </summary>
        private static string StripQuotes(
            string value)
        {
            return value
                .Trim()
                .Trim('"');
        }

        /// <summary>
        /// 根据常见 Office 路径做最后兜底。
        ///
        /// 注意:
        /// ClickToRun 环境下不应该依赖这个方法作为主路径，
        /// 主路径仍然来自 App Paths / 注册信息。
        /// </summary>
        private static string? FindExecutableFallback(
            OfficeDefinition definition,
            string? version)
        {
            var candidates =
                new[]
                {
                    Path.Combine(
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.ProgramFilesX86),
                        "Microsoft Office",
                        "Root",
                        "Office16",
                        definition.ExeName),

                    Path.Combine(
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.ProgramFiles),
                        "Microsoft Office",
                        "Root",
                        "Office16",
                        definition.ExeName)
                };

            foreach (var path in candidates)
            {
                if (File.Exists(path))
                    return path;
            }

            return null;
        }
    }

    // ====================================================================
    // COM 后期绑定
    //
    // 不依赖:
    //   dynamic
    //   Microsoft.Office.Interop.*
    //
    // 只使用 IDispatch 完成:
    //   GetProperty
    //   InvokeMethod
    // ====================================================================

    internal static class ComLateBinding
    {
        private const int DISPATCH_METHOD = 0x0001;
        private const int DISPATCH_PROPERTYGET = 0x0002;

        [ComImport]
        [Guid("00020400-0000-0000-C000-000000000046")]
        [InterfaceType(
            ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDispatch
        {
            [PreserveSig]
            int GetTypeInfoCount(
                out uint pctinfo);

            [PreserveSig]
            int GetTypeInfo(
                uint iTInfo,
                int lcid,
                out IntPtr ppTInfo);

            [PreserveSig]
            int GetIDsOfNames(
                ref Guid riid,
                [MarshalAs(
                    UnmanagedType.LPArray,
                    ArraySubType = UnmanagedType.LPWStr,
                    SizeParamIndex = 2)]
                string[] rgszNames,
                uint cNames,
                int lcid,
                [Out]
                [MarshalAs(
                    UnmanagedType.LPArray,
                    SizeParamIndex = 2)]
                int[] rgDispId);

            [PreserveSig]
            int Invoke(
                int dispIdMember,
                ref Guid riid,
                int lcid,
                int wFlags,
                ref DISPPARAMS pDispParams,
                [MarshalAs(
                    UnmanagedType.Struct)]
                out object pVarResult,
                out EXCEPINFO pExcepInfo,
                out uint puArgErr);
        }

        [StructLayout(
            LayoutKind.Sequential)]
        private struct DISPPARAMS
        {
            public IntPtr rgvarg;
            public IntPtr rgdispidNamedArgs;
            public int cArgs;
            public int cNamedArgs;
        }

        [StructLayout(
            LayoutKind.Sequential)]
        private struct EXCEPINFO
        {
            public short wCode;
            public short wReserved;

            [MarshalAs(
                UnmanagedType.BStr)]
            public string? bstrSource;

            [MarshalAs(
                UnmanagedType.BStr)]
            public string? bstrDescription;

            [MarshalAs(
                UnmanagedType.BStr)]
            public string? bstrHelpFile;

            public int dwHelpContext;
            public IntPtr pvReserved;
            public IntPtr pfnDeferredFillIn;
            public int scode;
        }

        /// <summary>
        /// 获取 COM 属性。
        /// </summary>
        public static object? GetProperty(
            object comObject,
            string name)
        {
            return Invoke(
                comObject,
                name,
                DISPATCH_PROPERTYGET,
                null);
        }

        /// <summary>
        /// 调用 COM 方法。
        /// </summary>
        public static object? InvokeMethod(
            object comObject,
            string name,
            params object[] args)
        {
            return Invoke(
                comObject,
                name,
                DISPATCH_METHOD,
                args);
        }

        private static object? Invoke(
            object comObject,
            string name,
            int flags,
            object?[]? args)
        {
            if (comObject == null)
                throw new ArgumentNullException(
                    nameof(comObject));

            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException(
                    "成员名称不能为空。",
                    nameof(name));

            if (comObject is not IDispatch dispatch)
            {
                throw new InvalidOperationException(
                    "COM 对象未提供 IDispatch。");
            }

            Guid riid = Guid.Empty;

            var names =
                new[] { name };

            var dispIds =
                new int[1];

            int hr =
                dispatch.GetIDsOfNames(
                    ref riid,
                    names,
                    1,
                    0,
                    dispIds);

            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(
                    hr);
            }

            int argCount =
                args?.Length ?? 0;

            IntPtr pArgs =
                IntPtr.Zero;

            int variantSize =
                Marshal.SizeOf<
                    VariantStruct>();

            try
            {
                var dispParams =
                    new DISPPARAMS
                    {
                        cArgs = 0,
                        cNamedArgs = 0,
                        rgvarg = IntPtr.Zero,
                        rgdispidNamedArgs =
                            IntPtr.Zero
                    };

                if (argCount > 0)
                {
                    pArgs =
                        Marshal.AllocHGlobal(
                            checked(
                                variantSize *
                                argCount));

                    // COM IDispatch 参数是反向排列。
                    for (int i = 0;
                         i < argCount;
                         i++)
                    {
                        var slot =
                            IntPtr.Add(
                                pArgs,
                                variantSize *
                                (argCount - 1 - i));

                        Marshal.GetNativeVariantForObject(
                            args![i],
                            slot);
                    }

                    dispParams.rgvarg =
                        pArgs;

                    dispParams.cArgs =
                        argCount;
                }

                EXCEPINFO excepInfo;

                object result;

                uint argError;

                hr =
                    dispatch.Invoke(
                        dispIds[0],
                        ref riid,
                        0,
                        flags,
                        ref dispParams,
                        out result,
                        out excepInfo,
                        out argError);

                if (hr < 0)
                {
                    if (!string.IsNullOrWhiteSpace(
                            excepInfo.bstrDescription))
                    {
                        throw new COMException(
                            excepInfo.bstrDescription,
                            hr);
                    }

                    Marshal.ThrowExceptionForHR(
                        hr);
                }

                return result;
            }
            finally
            {
                if (pArgs != IntPtr.Zero)
                {
                    for (int i = 0;
                         i < argCount;
                         i++)
                    {
                        var slot =
                            IntPtr.Add(
                                pArgs,
                                variantSize * i);

                        try
                        {
                            VariantClear(
                                slot);
                        }
                        catch
                        {
                            // Variant 清理失败不应覆盖原始异常。
                        }
                    }

                    Marshal.FreeHGlobal(
                        pArgs);
                }
            }
        }

        /// <summary>
        /// 仅用于计算原生 VARIANT 尺寸。
        ///
        /// x86:
        ///   16 bytes
        ///
        /// x64:
        ///   24 bytes
        ///
        /// 不把这个结构当作完整 VARIANT ABI 定义。
        /// </summary>
        [StructLayout(
            LayoutKind.Sequential)]
        private struct VariantStruct
        {
            public ushort vt;
            public ushort r1;
            public ushort r2;
            public ushort r3;

            public IntPtr p1;
            public IntPtr p2;
        }

        [DllImport(
            "oleaut32.dll",
            PreserveSig = false)]
        private static extern void VariantClear(
            IntPtr pvarg);
    }
}
