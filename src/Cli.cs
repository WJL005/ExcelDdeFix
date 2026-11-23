using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelDdeApp
{
    // ── 命令行输出：winexe 程序从 cmd 运行时挂回父控制台 ────
    internal static class ConsoleOut
    {
        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);
        [DllImport("kernel32.dll")]
        static extern bool FreeConsole();
        const int ATTACH_PARENT_PROCESS = -1;

        static bool attached = false;
        static readonly StringBuilder FileBuffer = new StringBuilder();

        public static void Init()
        {
            try { attached = AttachConsole(ATTACH_PARENT_PROCESS); }
            catch { attached = false; }
        }

        public static void Line(string s)
        {
            string line = s ?? "";
            FileBuffer.AppendLine(line);
            if (attached)
            {
                try { Console.WriteLine(line); }
                catch { }
            }
        }

        public static void FlushToFile(string name)
        {
            try
            {
                string dir = Log.Dir;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, name + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt");
                File.WriteAllText(path, FileBuffer.ToString(), new UTF8Encoding(false));
                Line("输出副本: " + path);
                // 再写一次，把上面这行也包含进去
                File.WriteAllText(path, FileBuffer.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }

        public static void Done()
        {
            if (attached) { try { FreeConsole(); } catch { } }
        }
    }

    // ── 实测合并验证：自动打开两个 xlsx，判断是否合并为单进程 ──
    // 关键安全设计：只统计并关闭"本次新拉起"的 Excel 进程，绝不碰用户已开的 Excel。
    public class VerifyResult
    {
        public bool Ok;
        public int NewExcelProcesses;
        public int NewExcelWindows;
        public List<int> NewPids = new List<int>();
        public string Message = "";
        public string File1 = "";
        public string File2 = "";
        public DateTime At;
    }

    internal static class MergeVerifier
    {
        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

        public static int CountXlsxWindows()
        {
            int n = 0;
            try
            {
                EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
                {
                    if (!IsWindowVisible(hWnd)) return true;
                    var sb = new StringBuilder(512);
                    GetWindowText(hWnd, sb, sb.Capacity);
                    string title = sb.ToString();
                    if (title.IndexOf(".xlsx", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        title.IndexOf(".xls", StringComparison.OrdinalIgnoreCase) >= 0) n++;
                    return true;
                }, IntPtr.Zero);
            }
            catch { }
            return n;
        }

        public static List<int> ExcelPids()
        {
            var list = new List<int>();
            try
            {
                foreach (Process p in Process.GetProcessesByName("EXCEL")) list.Add(p.Id);
            }
            catch { }
            return list;
        }

        public static VerifyResult Run(string f1, string f2, int waitMs, Action<string> progress)
        {
            var r = new VerifyResult();
            r.At = DateTime.Now;
            r.File1 = f1; r.File2 = f2;
            List<int> before = ExcelPids();

            if (progress != null) progress("记录打开前 Excel 进程: " + before.Count + " 个");
            try
            {
                Process.Start(new ProcessStartInfo(f1) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                r.Message = "打开第一个文件失败: " + ex.Message;
                return r;
            }
            System.Threading.Thread.Sleep(waitMs);
            if (progress != null) progress("已打开 " + Path.GetFileName(f1));

            try
            {
                Process.Start(new ProcessStartInfo(f2) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                r.Message = "打开第二个文件失败: " + ex.Message;
                Cleanup(before, r, progress);
                return r;
            }
            System.Threading.Thread.Sleep(waitMs + 1500);
            if (progress != null) progress("已打开 " + Path.GetFileName(f2));

            List<int> after = ExcelPids();
            foreach (int pid in after)
            {
                bool isNew = true;
                foreach (int old in before) if (old == pid) { isNew = false; break; }
                if (isNew) r.NewPids.Add(pid);
            }
            r.NewExcelProcesses = r.NewPids.Count;
            r.NewExcelWindows = CountXlsxWindows();
            r.Ok = (r.NewExcelProcesses == 1);

            if (r.Ok)
                r.Message = "合并成功：打开 2 个 xlsx 只新增 1 个 Excel 进程，DDE 合并生效。";
            else if (r.NewExcelProcesses == 0)
                r.Message = "未检测到新 Excel 进程，可能文件未真正打开（请检查文件关联）。";
            else
                r.Message = "未合并：打开 2 个 xlsx 新增了 " + r.NewExcelProcesses + " 个 Excel 进程（" +
                            string.Join(",", r.NewPids.ConvertAll(x => x.ToString()).ToArray()) +
                            "），DDE 合并未生效，复制粘贴可能丢格式。";

            if (progress != null) progress(r.Message);
            Cleanup(before, r, progress);
            return r;
        }

        // 只关闭本次新拉起的进程，且优先 CloseMainWindow（不粗暴杀死用户数据）
        static void Cleanup(List<int> before, VerifyResult r, Action<string> progress)
        {
            foreach (int pid in r.NewPids)
            {
                try
                {
                    Process p = Process.GetProcessById(pid);
                    p.CloseMainWindow();
                    if (progress != null) progress("请求关闭测试进程 PID=" + pid);
                }
                catch { }
            }
            System.Threading.Thread.Sleep(1500);
            foreach (int pid in r.NewPids)
            {
                try
                {
                    Process p = Process.GetProcessById(pid);
                    if (!p.HasExited)
                    {
                        p.Kill();
                        if (progress != null) progress("强制结束残留测试进程 PID=" + pid);
                    }
                }
                catch { }
            }
        }
    }

    // ── 命令行入口 ──────────────────────────────────────────
    // 退出码: 0=成功 1=部分失败/警告 2=参数错误 3=权限不足 4=异常
    internal static class Cli
    {
        static string Arg(string[] a, string name, string def = "")
        {
            for (int i = 0; i < a.Length; i++)
                if (string.Equals(a[i], name, StringComparison.OrdinalIgnoreCase) && i + 1 < a.Length)
                    return a[i + 1];
            return def;
        }

        static bool Has(string[] a, string name)
        {
            foreach (string s in a) if (string.Equals(s, name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        static string Json(string s)
        {
            if (s == null) return "\"\"";
            var sb = new StringBuilder("\"");
            foreach (char c in s)
            {
                if (c == '"') sb.Append("\\\"");
                else if (c == '\\') sb.Append("\\\\");
                else if (c == '\r') sb.Append("\\r");
                else if (c == '\n') sb.Append("\\n");
                else if (c < 32) sb.Append("\\u" + ((int)c).ToString("x4"));
                else sb.Append(c);
            }
            return sb.Append("\"").ToString();
        }

        public static int Run(string[] args)
        {
            ConsoleOut.Init();
            int rc = 0;
            try
            {
                if (args.Length == 0 || Has(args, "--help") || Has(args, "-h"))
                {
                    Help();
                    return 0;
                }
                if (Has(args, "--version") || Has(args, "-v"))
                {
                    ConsoleOut.Line("Excel DDE Fix v" + RegEngine.GetEnv().AppVersion);
                    return 0;
                }

                bool json = Has(args, "--json");
                string cmd = args[0];

                if (string.Equals(cmd, "--diagnose", StringComparison.OrdinalIgnoreCase))
                {
                    rc = Diagnose(json);
                }
                else if (string.Equals(cmd, "--backup", StringComparison.OrdinalIgnoreCase))
                {
                    rc = Backup(Arg(args, "--label", "CLI 备份"));
                }
                else if (string.Equals(cmd, "--repair", StringComparison.OrdinalIgnoreCase))
                {
                    rc = Repair(args);
                }
                else if (string.Equals(cmd, "--rollback", StringComparison.OrdinalIgnoreCase))
                {
                    rc = Rollback(args);
                }
                else if (string.Equals(cmd, "--verify", StringComparison.OrdinalIgnoreCase))
                {
                    rc = Verify(args);
                }
                else if (string.Equals(cmd, "--export", StringComparison.OrdinalIgnoreCase))
                {
                    rc = Export(Has(args, "--html"));
                }
                else if (string.Equals(cmd, "--logs", StringComparison.OrdinalIgnoreCase))
                {
                    int n = 30;
                    string ns = Arg(args, "--n", "30");
                    int.TryParse(ns, out n);
                    foreach (string l in Log.Tail(n)) ConsoleOut.Line(l);
                    rc = 0;
                }
                else
                {
                    ConsoleOut.Line("未知命令: " + cmd + "（用 --help 查看用法）");
                    rc = 2;
                }
            }
            catch (Exception ex)
            {
                Log.Exception("CLI", ex);
                ConsoleOut.Line("异常: " + ex.Message);
                rc = 4;
            }
            ConsoleOut.FlushToFile("cli");
            ConsoleOut.Done();
            return rc;
        }

        static void Help()
        {
            ConsoleOut.Line("Excel DDE 合并修复工具 v3 - 命令行用法");
            ConsoleOut.Line("");
            ConsoleOut.Line("  --diagnose [--json]        诊断当前状态（--json 输出机器可读 JSON）");
            ConsoleOut.Line("  --backup [--label 文字]    备份当前注册表状态");
            ConsoleOut.Line("  --repair [--dry-run] [--yes] 修复 DDE 合并（写前自动备份）");
            ConsoleOut.Line("                              --dry-run 只列出将写入什么，不改注册表");
            ConsoleOut.Line("                              --yes 跳过确认（用于自动化）");
            ConsoleOut.Line("  --rollback <备份文件> [--yes] 回滚到指定备份（先校验哈希）");
            ConsoleOut.Line("  --verify <文件1> <文件2>   实测：打开两个 xlsx 验证是否合并为单进程");
            ConsoleOut.Line("  --export [--html]          导出诊断报告（--html 导出可视化 HTML）");
            ConsoleOut.Line("  --logs [--n 行数]          查看最近日志");
            ConsoleOut.Line("  --version / --help         版本 / 本帮助");
            ConsoleOut.Line("");
            ConsoleOut.Line("退出码: 0=成功 1=部分失败/警告 2=参数错误 3=权限不足 4=异常");
        }

        static int Diagnose(bool json)
        {
            var h = RegEngine.Diagnose();
            var e = RegEngine.GetEnv();
            if (json)
            {
                var sb = new StringBuilder();
                sb.Append("{");
                sb.Append("\"version\":").Append(Json(e.AppVersion)).Append(",");
                sb.Append("\"score\":").Append(h.Score).Append(",");
                sb.Append("\"verdict\":").Append(Json(h.Verdict)).Append(",");
                sb.Append("\"excelExe\":").Append(Json(h.ExcelExe == null ? "" : h.ExcelExe)).Append(",");
                sb.Append("\"excelVersion\":").Append(Json(e.ExcelVersion)).Append(",");
                sb.Append("\"officeVersion\":").Append(Json(h.OfficeVersion)).Append(",");
                sb.Append("\"ddeCleaned\":").Append(Json(h.DdeCleaned)).Append(",");
                sb.Append("\"ddeCleanedKind\":").Append(Json(h.DdeCleanedKind)).Append(",");
                sb.Append("\"hardeningCount\":").Append(h.HardeningCount).Append(",");
                sb.Append("\"checks\":[");
                for (int i = 0; i < h.Checks.Count; i++)
                {
                    CheckItem c = h.Checks[i];
                    if (i > 0) sb.Append(",");
                    sb.Append("{").Append("\"progId\":").Append(Json(c.ProgId))
                      .Append(",\"item\":").Append(Json(c.Item))
                      .Append(",\"ok\":").Append(c.Ok ? "true" : "false")
                      .Append(",\"expected\":").Append(Json(c.Expected))
                      .Append(",\"actual\":").Append(Json(c.Actual)).Append("}");
                }
                sb.Append("]}");
                ConsoleOut.Line(sb.ToString());
            }
            else
            {
                ConsoleOut.Line(RegEngine.ExportReport());
            }
            return h.Score >= 90 ? 0 : 1;
        }

        static int Backup(string label)
        {
            var b = RegEngine.SaveBackup(label);
            ConsoleOut.Line("备份完成: " + b.FilePath);
            ConsoleOut.Line("条目=" + b.Values.Count + " 评分=" + b.Score + " 校验=" + b.Integrity);
            Log.WriteEx("BACKUP", "CLI 备份完成", "file=" + b.FilePath + " items=" + b.Values.Count + " score=" + b.Score);
            return 0;
        }

        static int Repair(string[] args)
        {
            // 演练模式不需要管理员权限：先让人看清楚"会改什么"，再决定是否提权
            string probe = RegEngine.FindExcelExe(true);
            if (Has(args, "--dry-run"))
            {
                ConsoleOut.Line("【演练模式】不会修改注册表。将执行：");
                ConsoleOut.Line("  探测到 Excel: " + (probe == null ? "(未找到，实际修复将中止)" : probe));
                ConsoleOut.Line("  command = \"" + (probe == null ? "EXCEL.EXE" : probe) + "\" /dde \"%1\"");
                ConsoleOut.Line("  ddeexec = " + DdeTargets.ExpectedDde);
                ConsoleOut.Line("  application = " + DdeTargets.ExpectedApp);
                ConsoleOut.Line("  topic = " + DdeTargets.ExpectedTopic);
                ConsoleOut.Line("  适用 ProgId : " + string.Join(", ", DdeTargets.ProgIds));
                ConsoleOut.Line("  当前评分: " + RegEngine.Diagnose().Score + "/100");
                ConsoleOut.Line("  注意：实际修复需要管理员权限（写 HKEY_CLASSES_ROOT），且会自动备份。");
                return 0;
            }

            bool admin = RegEngine.GetEnv().Admin;
            if (!admin)
            {
                ConsoleOut.Line("需要管理员权限才能写 HKEY_CLASSES_ROOT。请右键『以管理员身份运行』。");
                return 3;
            }
            string exe = probe;
            if (exe == null)
            {
                ConsoleOut.Line("无法定位 EXCEL.EXE（App Paths / COM / 已知路径都未找到）。已中止，避免写入错误路径。");
                return 1;
            }
            var before = RegEngine.CaptureCurrent();
            int scoreBefore = RegEngine.Diagnose().Score;

            if (!Has(args, "--yes"))
            {
                ConsoleOut.Line("确认修复？将自动备份并修改注册表。加 --yes 确认，或先用 --dry-run 演练。");
                return 2;
            }

            var b = RegEngine.SaveBackup("CLI 修复前自动备份");
            ConsoleOut.Line("自动备份: " + b.FilePath);
            var errs = RegEngine.Repair("dde", exe);
            var after = RegEngine.CaptureCurrent();
            int scoreAfter = RegEngine.Diagnose().Score;
            ConsoleOut.Line("评分: " + scoreBefore + " -> " + scoreAfter + "  失败项=" + errs.Count);
            foreach (string d in RegEngine.Diff(before, after)) ConsoleOut.Line("  " + d);
            foreach (string er in errs) ConsoleOut.Line("  ERROR " + er);
            Log.WriteEx("REPAIR", "CLI 修复完成", "score=" + scoreBefore + "->" + scoreAfter + " errs=" + errs.Count + " exe=" + exe);
            return errs.Count == 0 ? 0 : 1;
        }

        static int Rollback(string[] args)
        {
            string file = Arg(args, "--rollback", "");
            if (string.IsNullOrEmpty(file) || !File.Exists(file))
            {
                ConsoleOut.Line("用法: --rollback <备份文件路径> [--yes]");
                return 2;
            }
            if (!Has(args, "--yes"))
            {
                ConsoleOut.Line("确认回滚到 " + file + "？加 --yes 确认。");
                return 2;
            }
            var b = RegEngine.ParseBackup(file);
            string bad = RegEngine.ValidateBackup(b);
            if (bad != null) { ConsoleOut.Line("回滚被拒绝: " + bad); return 1; }
            ConsoleOut.Line("备份校验: " + b.Integrity + "  条目=" + b.Values.Count);
            var protect = RegEngine.SaveBackup("回滚前保护快照");
            ConsoleOut.Line("回滚前保护快照: " + protect.FilePath);
            int before = RegEngine.Diagnose().Score;
            var errs = RegEngine.Restore(b);
            int after = RegEngine.Diagnose().Score;
            ConsoleOut.Line("评分: " + before + " -> " + after + "  失败项=" + errs.Count);
            foreach (string er in errs) ConsoleOut.Line("  ERROR " + er);
            Log.WriteEx("ROLLBACK", "CLI 回滚完成", "file=" + file + " score=" + before + "->" + after + " errs=" + errs.Count);
            return errs.Count == 0 ? 0 : 1;
        }

        static int Verify(string[] args)
        {
            // 严格按 --verify <f1> <f2> 取两个位置参数，避免把 f1 误当成 f2
            string f1 = "", f2 = "";
            for (int i = 0; i < args.Length - 2; i++)
            {
                if (string.Equals(args[i], "--verify", StringComparison.OrdinalIgnoreCase))
                {
                    f1 = args[i + 1];
                    if (!args[i + 2].StartsWith("--")) f2 = args[i + 2];
                    break;
                }
            }
            if (string.IsNullOrEmpty(f1) || string.IsNullOrEmpty(f2) || !File.Exists(f1) || !File.Exists(f2))
            {
                ConsoleOut.Line("用法: --verify <文件1.xlsx> <文件2.xlsx>（两个文件都必须存在）");
                return 2;
            }
            var r = MergeVerifier.Run(f1, f2, 2500, delegate(string p)
            {
                ConsoleOut.Line("  " + p);
                Log.Write("VERIFY", p);   // 每一步都落盘，中途被中断也能看到进度
            });
            ConsoleOut.Line("实测结论: " + r.Message);
            Log.WriteEx("VERIFY", "CLI 实测合并验证", "ok=" + r.Ok + " newProcs=" + r.NewExcelProcesses +
                        " pids=" + string.Join(",", r.NewPids.ConvertAll(x => x.ToString()).ToArray()) +
                        " f1=" + r.File1 + " f2=" + r.File2);
            return r.Ok ? 0 : 1;
        }

        static int Export(bool html)
        {
            string ext = html ? "html" : "txt";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "diagnose-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + "." + ext);
            string content = html ? HtmlReport.Build() : RegEngine.ExportReport();
            File.WriteAllText(path, content, new UTF8Encoding(false));
            ConsoleOut.Line("报告已导出: " + path);
            Log.WriteEx("EXPORT", "CLI 导出报告", "file=" + path + " bytes=" + new FileInfo(path).Length);
            return 0;
        }
    }
}
