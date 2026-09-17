namespace OfficeChecker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnDetect_Click(object? sender, EventArgs e)
        {
            var sb = new System.Text.StringBuilder();
            btnDetect.Enabled = false;
            try
            {
                sb.AppendLine("进程位数: " + (Environment.Is64BitProcess ? "x64" : "x86"));
                sb.AppendLine();

                foreach (var kind in new[] { "Excel", "Word", "PowerPoint" })
                {
                    try
                    {
                        var report = OfficeProbe.OfficeDetector.DetectFull(kind);
                        var info = report.App;
                        sb.AppendLine($"[{kind}]");
                        sb.AppendLine($"  ProgID   : {info.ProgId}");
                        sb.AppendLine($"  CurVer   : {info.CurVer ?? "(未注册)"}");
                        sb.AppendLine($"  路径     : {info.ExePath ?? "(未找到)"}  存在: {info.ExeExists}");
                        sb.AppendLine($"  版本     : {info.Version ?? "?"}   内部版本: {info.Build ?? "?"}");
                        sb.AppendLine($"  文件版本 : {info.FileVersion ?? "?"}   产品版本: {info.ProductVersion ?? "?"}");
                        sb.AppendLine($"  渠道     : {info.Channel}   最终位数: {info.Bitness ?? "?"}");
                        sb.AppendLine($"  位数校验 : {report.Bitness.Message}");
                        if (info.BitnessConsistent == false)
                            sb.AppendLine("  !! 警告: 注册表位数与 PE 实际位数不一致，以 PE 为准");

                        var probe = report.Com;
                        sb.AppendLine($"  COM 实测 : {(probe.Ok ? "OK" : "FAIL")} {probe.Message}");
                        if (probe.Ok)
                        {
                            sb.AppendLine($"  Name     : {probe.DisplayName} {probe.Version}");
                            if (probe.DisplayName != null && !probe.DisplayName.StartsWith("Microsoft", StringComparison.OrdinalIgnoreCase))
                                sb.AppendLine("  !! 警告: ProgID 可能被 WPS 等第三方抢占，实际不是微软 Office");
                        }
                        else if (!string.IsNullOrWhiteSpace(probe.ExceptionMessage))
                        {
                            sb.AppendLine($"  异常     : {probe.ExceptionType} {probe.ExceptionMessage}");
                        }

                        sb.AppendLine($"  Interop  : {DescribeAssembly(report.InteropAssembly)}");
                        sb.AppendLine($"  NetOffice: {DescribeAssembly(report.NetOfficeAssembly)}");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"[{kind}] 检测异常: {ex.Message}");
                    }
                    sb.AppendLine();
                }

                txtResult.Text = sb.ToString();
            }
            finally
            {
                btnDetect.Enabled = true;
            }
        }

        private static string DescribeAssembly(OfficeProbe.AssemblyProbeResult? r)
        {
            if (r == null) return "(无)";
            return r.Status switch
            {
                OfficeProbe.AssemblyLoadStatus.Loaded => $"Loaded ({r.Location ?? r.FullName})",
                OfficeProbe.AssemblyLoadStatus.NotFound => $"NotFound {r.AssemblyName}",
                _ => $"LoadFailed {r.ErrorType}: {r.ErrorMessage}",
            };
        }
    }
}
