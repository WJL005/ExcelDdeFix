using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace ExcelDdeApp
{
    // ── 一条注册表项的快照值 ─────────────────────────────────
    public class RegValue
    {
        public string Path;      // 完整键路径，如 HKEY_CLASSES_ROOT\Excel.Sheet.12\shell\open\ddeexec
        public string Name;      // 值名，"" 表示默认值
        public string Value;     // 值内容；特殊标记 "<MISSING>" 表示该值/键原本不存在
        public string Kind;      // String / DWord / Missing —— 注册表原类型

        public RegValue(string path, string name, string value, string kind)
        {
            Path = path; Name = name; Value = value; Kind = kind;
        }

        public string Display
        {
            get { return string.IsNullOrEmpty(Name) ? "(默认)" : Name; }
        }

        public string Key { get { return Path + "\\" + Name; } }
    }

    // ── 一次完整备份 ────────────────────────────────────────
    public class DdeBackup
    {
        public DateTime At;
        public string Label;
        public List<RegValue> Values = new List<RegValue>();
        public string FilePath;
        public string Integrity = "v1(unverified)";   // v2 备份带 SHA256 校验
        public int Score = -1;                        // 备份时的健康评分
    }

    // ── 一项精确校验结果（v3 评分：校验值内容，不只是键存在）──
    public class CheckItem
    {
        public string ProgId;
        public string Item;        // ddeexec / application / topic / command
        public bool Ok;
        public string Expected;
        public string Actual;

        public CheckItem(string progId, string item, bool ok, string expected, string actual)
        {
            ProgId = progId; Item = item; Ok = ok; Expected = expected; Actual = actual;
        }
    }

    // ── 环境指纹（开源排障必需：报告要能复现环境）──────────
    public class EnvInfo
    {
        public string AppVersion = "3.0.0";
        public string Machine = "";
        public string User = "";
        public string Os = "";
        public string OsBits = "";
        public string DotNet = "";
        public string ExcelPath = "";
        public string ExcelVersion = "";
        public string ExcelBits = "";
        public bool Admin;
        public string OfficeVersionsFound = "";
        public string Now = "";
    }

    // ── 我们关心的所有注册表位置 ────────────────────────────
    public static class DdeTargets
    {
        public const string HKCR = "HKEY_CLASSES_ROOT";
        public const string HKCU = "HKEY_CURRENT_USER";
        public const string HKLM = "HKEY_LOCAL_MACHINE";

        // Excel ProgId（Word/PPT 仅体检不改）
        public static readonly string[] ProgIds = new string[]
        {
            "Excel.Sheet.12",                       // .xlsx
            "Excel.Sheet.8",                        // .xls
            "Excel.SheetMacroEnabled.12",           // .xlsm
            "Excel.SheetBinaryMacroEnabled.12",     // .xlsb
            "Word.Document.12",                     // .docx（仅体检）
            "PowerPoint.Show.12"                    // .pptx（仅体检）
        };

        public static bool IsExcel(string progId)
        {
            return progId != null && progId.StartsWith("Excel.", StringComparison.OrdinalIgnoreCase);
        }

        public static string CmdPath(string p)  { return HKCR + "\\" + p + "\\shell\\open\\command"; }
        public static string DdePath(string p)  { return HKCR + "\\" + p + "\\shell\\open\\ddeexec"; }
        public static string AppPath(string p)  { return HKCR + "\\" + p + "\\shell\\open\\ddeexec\\application"; }
        public static string TopicPath(string p){ return HKCR + "\\" + p + "\\shell\\open\\ddeexec\\topic"; }

        // ── 多版本 Office 支持（旧版只认 16.0）────────────────
        public static readonly string[] OfficeVersions = new string[] { "16.0", "15.0", "14.0" };

        public static string ExcelOptions(string ver)
        {
            return HKCU + "\\Software\\Microsoft\\Office\\" + ver + "\\Excel\\Options";
        }

        // 返回本机实际存在的 Office 版本列表（用于备份/报告全覆盖）
        public static string[] InstalledOfficeVersions()
        {
            var list = new List<string>();
            foreach (string v in OfficeVersions)
            {
                string path = ExcelOptions(v);
                try
                {
                    using (RegistryKey k = RegEngine.OpenKeyPublic(path, false))
                    {
                        if (k != null) list.Add(v);
                    }
                }
                catch { }
            }
            return list.ToArray();
        }

        // 主版本：优先 16.0，否则取第一个存在的；都不存在则 16.0
        public static string PrimaryOfficeVersion()
        {
            string[] vs = InstalledOfficeVersions();
            if (vs.Length == 0) return "16.0";
            foreach (string v in vs) if (v == "16.0") return v;
            return vs[0];
        }

        // 兼容旧引用的主版本 Options 路径（动态解析）
        public static string ExcelOptionsPath { get { return ExcelOptions(PrimaryOfficeVersion()); } }

        // ── 修复期望值（v3 精确校验的基准）────────────────────
        public const string ExpectedDde       = "[open(\"%1\")]";
        public const string ExpectedApp       = "EXCEL.EXE";
        public const string ExpectedTopic     = "System";
        public const string ExpectedDdeSwitch = "/dde";

        // ── DDE 安全加固监测项 ───────────────────────────────
        // 这些值是安全加固工具（HardeningKitty / DDEAuto 防护脚本 / 组策略）写入的，
        // 目的：防 DDE 攻击。它们不是"故障"，独立监测、不计入合并评分。
        public const string S_DDECleaned   = "DDECleaned";
        public const string S_DDEAllowed   = "DDEAllowed";
        public const string S_DontUpdateLinks = "DontUpdateLinks";
        public const string S_IgnoreOtherApps = "IgnoreOtherApplications";

        public const string PolicyExtContent = HKCU + "\\Software\\Policies\\Microsoft\\Office\\16.0\\excel\\security\\external content";
        public const string PolicyExtContentLM = HKLM + "\\SOFTWARE\\Policies\\Microsoft\\Office\\16.0\\excel\\security\\external content";
        public const string S_DisableServerLaunch = "DisableDdeServerLaunch";
        public const string S_DisableServerLookup = "DisableDdeServerLookup";

        public class SafetyTarget
        {
            public string Path;
            public string Name;
            public string Desc;
            public SafetyTarget(string path, string name, string desc)
            { Path = path; Name = name; Desc = desc; }
        }

        // 主版本安全项（界面展示用，8 项）
        public static SafetyTarget[] SafetyTargets
        {
            get
            {
                string opt = ExcelOptions(PrimaryOfficeVersion());
                return new SafetyTarget[]
                {
                    new SafetyTarget(opt, S_DDECleaned,     "已做 DDE 安全加固标记（1=已加固，不影响合并）"),
                    new SafetyTarget(opt, S_DDEAllowed,     "DDE 总开关（0=禁用 DDE 加固；未设置=默认允许）"),
                    new SafetyTarget(opt, S_DontUpdateLinks,"禁止自动更新工作簿链接（1=禁止，安全加固）"),
                    new SafetyTarget(opt, S_IgnoreOtherApps,"忽略其他程序 DDE 请求（1=忽略）"),
                    new SafetyTarget(PolicyExtContent,    S_DisableServerLaunch, "策略：禁止 DDE 服务器启动（1=禁止）"),
                    new SafetyTarget(PolicyExtContent,    S_DisableServerLookup, "策略：禁止 DDE 服务器查找（1=禁止）"),
                    new SafetyTarget(PolicyExtContentLM,  S_DisableServerLaunch, "策略(机器)：禁止 DDE 服务器启动（1=禁止）"),
                    new SafetyTarget(PolicyExtContentLM,  S_DisableServerLookup, "策略(机器)：禁止 DDE 服务器查找（1=禁止）")
                };
            }
        }

        // 全版本安全项（备份/报告用，覆盖所有已安装 Office 版本）
        public static SafetyTarget[] AllSafetyTargets()
        {
            var list = new List<SafetyTarget>();
            foreach (string v in InstalledOfficeVersions())
            {
                string opt = ExcelOptions(v);
                list.Add(new SafetyTarget(opt, S_DDECleaned,      "[" + v + "] 已做 DDE 安全加固标记"));
                list.Add(new SafetyTarget(opt, S_DDEAllowed,      "[" + v + "] DDE 总开关"));
                list.Add(new SafetyTarget(opt, S_DontUpdateLinks, "[" + v + "] 禁止自动更新工作簿链接"));
                list.Add(new SafetyTarget(opt, S_IgnoreOtherApps, "[" + v + "] 忽略其他程序 DDE 请求"));
            }
            list.Add(new SafetyTarget(PolicyExtContent,    S_DisableServerLaunch, "策略：禁止 DDE 服务器启动"));
            list.Add(new SafetyTarget(PolicyExtContent,    S_DisableServerLookup, "策略：禁止 DDE 服务器查找"));
            list.Add(new SafetyTarget(PolicyExtContentLM,  S_DisableServerLaunch, "策略(机器)：禁止 DDE 服务器启动"));
            list.Add(new SafetyTarget(PolicyExtContentLM,  S_DisableServerLookup, "策略(机器)：禁止 DDE 服务器查找"));
            return list.ToArray();
        }
    }

    public static class RegEngine
    {
        static readonly object Gate = new object();

        // 把 "HKEY_CLASSES_ROOT\..." 转成 RegistryKey（公开供 DdeTargets 探测用）
        public static RegistryKey OpenKeyPublic(string fullPath, bool writable)
        {
            return OpenKey(fullPath, writable);
        }

        static RegistryKey OpenKey(string fullPath, bool writable)
        {
            if (string.IsNullOrEmpty(fullPath)) return null;
            string[] parts = fullPath.Split('\\');
            RegistryKey root = GetRoot(parts[0]);
            if (root == null) return null;
            string sub = string.Join("\\", parts, 1, parts.Length - 1);
            return writable ? root.OpenSubKey(sub, true) : root.OpenSubKey(sub, false);
        }

        static RegistryKey GetRoot(string name)
        {
            if (name == "HKEY_CLASSES_ROOT") return Registry.ClassesRoot;
            if (name == "HKEY_CURRENT_USER") return Registry.CurrentUser;
            if (name == "HKEY_LOCAL_MACHINE") return Registry.LocalMachine;
            if (name == "HKEY_USERS") return Registry.Users;
            return null;
        }

        // 读取值；同时通过 out 返回注册表原类型
        public static string ReadValue(string fullPath, string name, out string kind)
        {
            kind = "Missing";
            try
            {
                using (RegistryKey k = OpenKey(fullPath, false))
                {
                    if (k == null) return "<MISSING>";
                    object v = k.GetValue(name, null);
                    if (v == null) return "<MISSING>";
                    RegistryValueKind vk;
                    try { vk = k.GetValueKind(name); }
                    catch { vk = RegistryValueKind.String; }
                    if (vk == RegistryValueKind.DWord) kind = "DWord";
                    else if (vk == RegistryValueKind.String || vk == RegistryValueKind.ExpandString) kind = "String";
                    else kind = vk.ToString();
                    return v.ToString();
                }
            }
            catch { return "<MISSING>"; }
        }

        public static string ReadValue(string fullPath, string name)
        {
            string kind;
            return ReadValue(fullPath, name, out kind);
        }

        public static bool KeyExists(string fullPath)
        {
            try { using (RegistryKey k = OpenKey(fullPath, false)) return k != null; }
            catch { return false; }
        }

        // 按指定类型写入（修正旧工具一律写 String 的 bug）
        public static bool WriteValue(string fullPath, string name, string value, string kind)
        {
            try
            {
                string[] parts = fullPath.Split('\\');
                RegistryKey root = GetRoot(parts[0]);
                if (root == null) return false;
                string sub = string.Join("\\", parts, 1, parts.Length - 1);
                using (RegistryKey k = root.CreateSubKey(sub, true))
                {
                    if (k == null) return false;
                    if (kind == "DWord")
                    {
                        int n;
                        if (!int.TryParse(value, out n)) n = 0;
                        k.SetValue(name, n, RegistryValueKind.DWord);
                    }
                    else
                    {
                        k.SetValue(name, value, RegistryValueKind.String);
                    }
                    return true;
                }
            }
            catch { return false; }
        }

        public static bool WriteValue(string fullPath, string name, string value)
        {
            return WriteValue(fullPath, name, value, "String");
        }

        public static bool DeleteTree(string fullPath)
        {
            try
            {
                string[] parts = fullPath.Split('\\');
                RegistryKey root = GetRoot(parts[0]);
                if (root == null) return false;
                string sub = string.Join("\\", parts, 1, parts.Length - 1);
                root.DeleteSubKeyTree(sub, false);
                return true;
            }
            catch { return false; }
        }

        public static bool DeleteValue(string fullPath, string name)
        {
            try
            {
                using (RegistryKey k = OpenKey(fullPath, true))
                {
                    if (k == null) return true; // 本来就不存在，视为成功
                    k.DeleteValue(name, false);
                    return true;
                }
            }
            catch { return false; }
        }

        // ── 读取当前所有相关项，形成"值清单"（含类型）──────────
        public static List<RegValue> CaptureCurrent()
        {
            var list = new List<RegValue>();
            foreach (string p in DdeTargets.ProgIds)
            {
                AddCaptured(list, DdeTargets.CmdPath(p), "");
                if (DdeTargets.IsExcel(p))
                {
                    AddCaptured(list, DdeTargets.DdePath(p), "");
                    AddCaptured(list, DdeTargets.AppPath(p), "");
                    AddCaptured(list, DdeTargets.TopicPath(p), "");
                }
            }
            // 安全加固项也纳入备份（含类型），回滚才能完整还原
            foreach (var t in DdeTargets.AllSafetyTargets())
                AddCaptured(list, t.Path, t.Name);
            return list;
        }

        static void AddCaptured(List<RegValue> list, string path, string name)
        {
            string kind;
            string val = ReadValue(path, name, out kind);
            list.Add(new RegValue(path, name, val, val == "<MISSING>" ? "Missing" : kind));
        }

        // ── 精确差异：逐项对比两个快照 ───────────────────────
        public static List<string> Diff(List<RegValue> before, List<RegValue> after)
        {
            var res = new List<string>();
            var map = new Dictionary<string, RegValue>(StringComparer.OrdinalIgnoreCase);
            if (before != null) foreach (var v in before) map[v.Key] = v;
            if (after == null) return res;
            foreach (var v in after)
            {
                RegValue old;
                if (!map.TryGetValue(v.Key, out old))
                {
                    res.Add("+ " + ShortKey(v.Key) + " = " + Clip(v.Value) + " [" + v.Kind + "]");
                    continue;
                }
                if (old.Value != v.Value || old.Kind != v.Kind)
                {
                    res.Add("~ " + ShortKey(v.Key) + ": " + Clip(old.Value) + " [" + old.Kind + "]" +
                            "  ->  " + Clip(v.Value) + " [" + v.Kind + "]");
                }
            }
            foreach (var v in before)
            {
                bool found = false;
                foreach (var a in after) if (string.Equals(a.Key, v.Key, StringComparison.OrdinalIgnoreCase)) { found = true; break; }
                if (!found) res.Add("- " + ShortKey(v.Key) + " (原本 " + Clip(v.Value) + ")");
            }
            return res;
        }

        static string ShortKey(string key)
        {
            string s = key.Replace("HKEY_CLASSES_ROOT\\", "HKCR\\")
                          .Replace("HKEY_CURRENT_USER\\", "HKCU\\")
                          .Replace("HKEY_LOCAL_MACHINE\\", "HKLM\\");
            int i = s.LastIndexOf("\\", StringComparison.Ordinal);
            if (i > 0 && i < s.Length - 1 && s.Substring(i + 1).Length <= 24)
                return s;
            return s;
        }

        static string Clip(string s)
        {
            if (s == null) return "(null)";
            return s.Length > 44 ? s.Substring(0, 44) + "..." : s;
        }

        // ── 备份目录 ────────────────────────────────────────
        public static string BackupDir
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "backup"); }
        }

        public static DdeBackup SaveBackup(string label)
        {
            lock (Gate)
            {
                if (!Directory.Exists(BackupDir)) Directory.CreateDirectory(BackupDir);
                var b = new DdeBackup();
                b.At = DateTime.Now;
                b.Label = label;
                b.Values = CaptureCurrent();
                b.Score = Diagnose().Score;
                string stamp = b.At.ToString("yyyyMMdd-HHmmss");
                b.FilePath = Path.Combine(BackupDir, "dde-backup-" + stamp + ".txt");

                var sb = new StringBuilder();
                sb.AppendLine("# Excel DDE 修复工具 - 状态备份 v3");
                sb.AppendLine("at=" + b.At.ToString("yyyy-MM-dd HH:mm:ss"));
                sb.AppendLine("label=" + label);
                sb.AppendLine("score=" + b.Score);
                sb.AppendLine("# 格式: PATH<TAB>VALUENAME<TAB>KIND<TAB>VALUE  (KIND=Missing 表示原本不存在)");
                foreach (RegValue v in b.Values)
                    sb.AppendLine(v.Path + "\t" + v.Name + "\t" + v.Kind + "\t" + v.Value);

                // v3：内容哈希，防止备份被误改/损坏后仍被回滚
                string hash = Sha256(sb.ToString());
                sb.AppendLine("# sha256=" + hash);
                File.WriteAllText(b.FilePath, sb.ToString(), new UTF8Encoding(false));
                b.Integrity = "v3(sha256=" + hash.Substring(0, 12) + ")";
                return b;
            }
        }

        public static string Sha256(string text)
        {
            try
            {
                byte[] data = new UTF8Encoding(false).GetBytes(text);
                using (SHA256 sha = SHA256.Create())
                {
                    byte[] h = sha.ComputeHash(data);
                    var sb = new StringBuilder();
                    foreach (byte x in h) sb.Append(x.ToString("x2"));
                    return sb.ToString();
                }
            }
            catch { return ""; }
        }

        public static List<DdeBackup> ListBackups()
        {
            var list = new List<DdeBackup>();
            try
            {
                if (!Directory.Exists(BackupDir)) return list;
                foreach (string f in Directory.GetFiles(BackupDir, "dde-backup-*.txt"))
                {
                    var b = ParseBackup(f);
                    if (b != null) list.Add(b);
                }
                list.Sort((a, c) => c.At.CompareTo(a.At));
            }
            catch { }
            return list;
        }

        public static DdeBackup ParseBackup(string file)
        {
            try
            {
                var b = new DdeBackup();
                b.FilePath = file;
                b.At = File.GetLastWriteTime(file);
                string[] lines = File.ReadAllLines(file, new UTF8Encoding(false));
                string expectedHash = "";
                var content = new StringBuilder();

                foreach (string raw in lines)
                {
                    if (raw.StartsWith("# sha256=", StringComparison.OrdinalIgnoreCase))
                    {
                        expectedHash = raw.Substring(9).Trim();
                        continue;
                    }
                    if (raw.StartsWith("#") || string.IsNullOrWhiteSpace(raw)) continue;
                    content.AppendLine(raw);
                    if (raw.StartsWith("at="))
                    {
                        DateTime dt;
                        if (DateTime.TryParse(raw.Substring(3), out dt)) b.At = dt;
                        continue;
                    }
                    if (raw.StartsWith("label=")) { b.Label = raw.Substring(6); continue; }
                    if (raw.StartsWith("score="))
                    {
                        int sc;
                        if (int.TryParse(raw.Substring(6), out sc)) b.Score = sc;
                        continue;
                    }
                    string[] parts = raw.Split('\t');
                    if (parts.Length >= 4)
                        b.Values.Add(new RegValue(parts[0], parts[1], parts[3], parts[2]));
                    else if (parts.Length == 3)
                        b.Values.Add(new RegValue(parts[0], parts[1], parts[2],
                            parts[2] == "<MISSING>" ? "Missing" : "String")); // 兼容旧 v1 备份
                }

                if (!string.IsNullOrEmpty(expectedHash))
                {
                    string actual = Sha256(content.ToString());
                    b.Integrity = (actual == expectedHash)
                        ? "v3(已校验 OK)"
                        : "v3(校验失败! 文件可能被修改)";
                }
                return b;
            }
            catch { return null; }
        }

        // 回滚前的安全检查：返回 null 表示可用，否则返回拒绝原因
        public static string ValidateBackup(DdeBackup b)
        {
            if (b == null) return "备份为空";
            if (b.Values.Count == 0) return "备份没有任何条目";
            if (b.Integrity.IndexOf("校验失败", StringComparison.Ordinal) >= 0)
                return "备份文件哈希校验失败，可能已被修改或损坏，拒绝回滚";
            // 必须有 command 项，否则说明备份内容不完整
            bool hasCmd = false;
            foreach (var v in b.Values)
                if (v.Path.EndsWith("\\shell\\open\\command", StringComparison.OrdinalIgnoreCase)) { hasCmd = true; break; }
            if (!hasCmd) return "备份缺少 command 项，内容不完整，拒绝回滚";
            return null;
        }

        // ── 回滚：把备份里的每个值按原类型写回去 ────────────
        // 安全护栏：只允许删除 \ddeexec 结尾的键树，防止恶意/损坏备份删掉 command 键
        public static List<string> Restore(DdeBackup b)
        {
            var errs = new List<string>();
            if (b == null) { errs.Add("备份为空"); return errs; }
            string invalid = ValidateBackup(b);
            if (invalid != null) { errs.Add("回滚被拒绝: " + invalid); return errs; }
            lock (Gate)
            {
                foreach (RegValue v in b.Values)
                {
                    bool ok;
                    if (v.Kind == "Missing")
                    {
                        if (string.IsNullOrEmpty(v.Name))
                        {
                            if (v.Path.EndsWith("\\ddeexec", StringComparison.OrdinalIgnoreCase))
                                ok = DeleteTree(v.Path);          // 仅允许删 ddeexec 键树
                            else
                            {
                                errs.Add("跳过危险的整键删除: " + v.Path);
                                continue;
                            }
                        }
                        else
                            ok = DeleteValue(v.Path, v.Name);     // 值原本不存在 -> 删除值
                    }
                    else
                    {
                        ok = WriteValue(v.Path, v.Name, v.Value, v.Kind); // 按原类型还原
                    }
                    if (!ok) errs.Add(v.Path + "\\" + v.Display + " [" + v.Kind + "]");
                }
            }
            return errs;
        }

        // ── 修复：恢复 DDE 合并 ─────────────────────────────
        // mode: "dde"    = 补 ddeexec + command 加 /dde（推荐，不动安全加固语义）
        //       "clean"  = 仅把被写坏的 DDECleaned 类型修正回标准 DWord 1
        public static List<string> Repair(string mode, string excelExe)
        {
            var errs = new List<string>();
            lock (Gate)
            {
                if (mode == "clean")
                {
                    string opt = DdeTargets.ExcelOptions(DdeTargets.PrimaryOfficeVersion());
                    if (!WriteValue(opt, DdeTargets.S_DDECleaned, "1", "DWord"))
                        errs.Add("修复 DDECleaned 类型失败");
                    return errs;
                }

                foreach (string p in DdeTargets.ProgIds)
                {
                    if (!DdeTargets.IsExcel(p)) continue; // Word/PPT 只体检不改

                    string cmd = "\"" + excelExe + "\" /dde \"%1\"";
                    if (!WriteValue(DdeTargets.CmdPath(p), "", cmd))
                        errs.Add("写 command 失败: " + p);
                    if (!WriteValue(DdeTargets.DdePath(p), "", DdeTargets.ExpectedDde))
                        errs.Add("写 ddeexec 失败: " + p);
                    if (!WriteValue(DdeTargets.AppPath(p), "", DdeTargets.ExpectedApp))
                        errs.Add("写 ddeexec\\application 失败: " + p);
                    if (!WriteValue(DdeTargets.TopicPath(p), "", DdeTargets.ExpectedTopic))
                        errs.Add("写 ddeexec\\topic 失败: " + p);
                }

                // 顺手修正被旧工具写坏的 DDECleaned 类型（String -> DWord 1，保持加固语义）
                string opt2 = DdeTargets.ExcelOptions(DdeTargets.PrimaryOfficeVersion());
                string ddk;
                string ddc = ReadValue(opt2, DdeTargets.S_DDECleaned, out ddk);
                if (ddc != "<MISSING>" && ddk != "DWord")
                    if (!WriteValue(opt2, DdeTargets.S_DDECleaned, "1", "DWord"))
                        errs.Add("修正 DDECleaned 类型失败");
            }
            return errs;
        }

        // ── Excel 路径智能探测（旧版硬编码 4 个路径，会写错路径）──
        // 优先级：App Paths(注册表) > COM LocalServer32 > 已知路径
        public static string FindExcelExe(bool requireExists = true)
        {
            var candidates = new List<string>();

            // 1) App Paths（Windows 官方的"按名字找程序"入口）
            string ap1 = ReadValue(DdeTargets.HKLM + "\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe", "");
            string ap2 = ReadValue(DdeTargets.HKCU + "\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe", "");
            if (!string.IsNullOrEmpty(ap1)) candidates.Add(ExtractExePath(ap1));
            if (!string.IsNullOrEmpty(ap2)) candidates.Add(ExtractExePath(ap2));

            // 2) COM 自动化注册：Excel.Application\CLSID -> LocalServer32
            string clsid = ReadValue(DdeTargets.HKCR + "\\Excel.Application\\CLSID", "");
            if (!string.IsNullOrEmpty(clsid))
            {
                string ls = ReadValue(DdeTargets.HKCR + "\\CLSID\\" + clsid + "\\LocalServer32", "");
                if (!string.IsNullOrEmpty(ls)) candidates.Add(ExtractExePath(ls));
            }

            // 3) 已知路径：多版本 Office（2010/2013/2016+/C2R）+ 32/64 位
            string[] vers = new string[] { "Root\\Office16", "Office16", "Root\\Office15", "Office15", "Root\\Office14", "Office14" };
            string[] pf = new string[]
            {
                "C:\\Program Files\\Microsoft Office\\",
                "C:\\Program Files (x86)\\Microsoft Office\\",
                "D:\\Program Files\\Microsoft Office\\",
                "D:\\Program Files (x86)\\Microsoft Office\\"
            };
            foreach (string b in pf)
                foreach (string v in vers)
                    candidates.Add(b + v + "\\EXCEL.EXE");

            foreach (string c in candidates)
            {
                if (string.IsNullOrEmpty(c)) continue;
                if (!c.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) continue;
                if (!requireExists || File.Exists(c)) return c;
            }
            return requireExists ? null : candidates[0];
        }

        // 从 "C:\path\EXCEL.EXE" /automation 中提取纯路径
        public static string ExtractExePath(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return raw;
            string s = raw.Trim();
            if (s.StartsWith("\""))
            {
                int end = s.IndexOf('"', 1);
                if (end > 1) return s.Substring(1, end - 1);
            }
            int idx = s.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0) return s.Substring(0, idx + 4);
            return s;
        }

        // ── 环境指纹 ────────────────────────────────────────
        public static EnvInfo GetEnv()
        {
            var e = new EnvInfo();
            e.Machine = Environment.MachineName;
            e.User = Environment.UserName;
            e.Os = Environment.OSVersion.VersionString;
            e.OsBits = IntPtr.Size == 8 ? "64-bit" : "32-bit";
            e.DotNet = Environment.Version.ToString();
            e.Now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            try
            {
                e.Admin = new System.Security.Principal.WindowsPrincipal(
                    System.Security.Principal.WindowsIdentity.GetCurrent()).IsInRole(
                    System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch { }
            string exe = FindExcelExe(true);
            e.ExcelPath = exe == null ? "(未找到)" : exe;
            if (exe != null)
            {
                try
                {
                    FileVersionInfo vi = FileVersionInfo.GetVersionInfo(exe);
                    e.ExcelVersion = vi.ProductVersion;
                }
                catch { }
                e.ExcelBits = GetPeArchitecture(exe);
            }
            e.OfficeVersionsFound = string.Join(",", DdeTargets.InstalledOfficeVersions());
            return e;
        }

        // 读 PE 头判断 x86/x64（比"看路径里有没有 (x86)"可靠）
        public static string GetPeArchitecture(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (BinaryReader br = new BinaryReader(fs))
                {
                    if (fs.Length < 0x40) return "未知";
                    fs.Position = 0x3C;
                    int peOffset = br.ReadInt32();
                    if (peOffset <= 0 || peOffset + 6 > fs.Length) return "未知";
                    fs.Position = peOffset;
                    if (br.ReadUInt32() != 0x00004550) return "未知"; // "PE\0\0"
                    ushort machine = br.ReadUInt16();
                    if (machine == 0x8664) return "64-bit";
                    if (machine == 0x014C) return "32-bit";
                    if (machine == 0x0200) return "IA64";
                    return "未知(0x" + machine.ToString("X4") + ")";
                }
            }
            catch { return "未知"; }
        }

        // ── 健康诊断（v3：精确值校验 + 4×4 项各 6.25 分）────
        public class HealthInfo
        {
            public int ExcelProgIds;
            public int DdeOk;
            public int CmdWithDde;
            public string DdeCleaned = "<MISSING>";
            public string DdeCleanedKind = "Missing";
            public string ExcelProcesses;
            public string OpenWorkbooks;
            public int Score;
            public string Verdict;
            public string DdeAllowed = "<MISSING>";
            public string DontUpdateLinks = "<MISSING>";
            public string IgnoreOtherApps = "<MISSING>";
            public int HardeningCount;
            public List<string> SafetyNotes = new List<string>();
            public List<CheckItem> Checks = new List<CheckItem>();
            public string ExcelExe = "";
            public string OfficeVersion = "";
        }

        static bool Same(string a, string b)
        {
            return string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
        }

        public static HealthInfo Diagnose()
        {
            var h = new HealthInfo();
            int total = 0, dde = 0, cmd = 0;
            h.OfficeVersion = DdeTargets.PrimaryOfficeVersion();
            h.ExcelExe = FindExcelExe(true);
            string exeName = h.ExcelExe == null ? "EXCEL.EXE" : Path.GetFileName(h.ExcelExe);

            foreach (string p in DdeTargets.ProgIds)
            {
                if (!DdeTargets.IsExcel(p)) continue;
                total++;

                // 1) ddeexec 默认值必须等于 [open("%1")]
                string ddeVal = ReadValue(DdeTargets.DdePath(p), "");
                bool ddeExists = KeyExists(DdeTargets.DdePath(p));
                bool ddeOk = ddeExists && Same(ddeVal, DdeTargets.ExpectedDde);
                h.Checks.Add(new CheckItem(p, "ddeexec", ddeOk, DdeTargets.ExpectedDde, ddeVal));
                if (ddeOk) dde++;

                // 2) application 必须指向 EXCEL.EXE（允许全路径或小写）
                string appVal = ReadValue(DdeTargets.AppPath(p), "");
                bool appOk = appVal.EndsWith(DdeTargets.ExpectedApp, StringComparison.OrdinalIgnoreCase);
                h.Checks.Add(new CheckItem(p, "application", appOk, DdeTargets.ExpectedApp, appVal));

                // 3) topic 必须是 System
                string topicVal = ReadValue(DdeTargets.TopicPath(p), "");
                bool topicOk = Same(topicVal, DdeTargets.ExpectedTopic);
                h.Checks.Add(new CheckItem(p, "topic", topicOk, DdeTargets.ExpectedTopic, topicVal));

                // 4) command 必须含 /dde 且指向存在的 Excel
                string cmdVal = ReadValue(DdeTargets.CmdPath(p), "");
                bool cmdOk = cmdVal.IndexOf(DdeTargets.ExpectedDdeSwitch, StringComparison.OrdinalIgnoreCase) >= 0 &&
                             cmdVal.IndexOf(exeName, StringComparison.OrdinalIgnoreCase) >= 0;
                h.Checks.Add(new CheckItem(p, "command", cmdOk, "\"" + exeName + "\" /dde \"%1\"", cmdVal));
                if (cmdOk) cmd++;
            }
            h.ExcelProgIds = total;
            h.DdeOk = dde;
            h.CmdWithDde = cmd;

            string opt = DdeTargets.ExcelOptions(h.OfficeVersion);
            h.DdeCleaned = ReadValue(opt, DdeTargets.S_DDECleaned, out h.DdeCleanedKind);
            h.DdeAllowed = ReadValue(opt, DdeTargets.S_DDEAllowed);
            h.DontUpdateLinks = ReadValue(opt, DdeTargets.S_DontUpdateLinks);
            h.IgnoreOtherApps = ReadValue(opt, DdeTargets.S_IgnoreOtherApps);
            string kind;
            foreach (var t in DdeTargets.SafetyTargets)
            {
                string v = ReadValue(t.Path, t.Name, out kind);
                if (v != "<MISSING>" && v != "0")
                {
                    h.HardeningCount++;
                    if (t.Name == DdeTargets.S_DisableServerLaunch && v == "1")
                        h.SafetyNotes.Add("策略禁止 DDE 服务器启动（实测不影响文件打开合并）");
                }
            }

            try
            {
                Process[] ps = Process.GetProcessesByName("EXCEL");
                h.ExcelProcesses = ps.Length.ToString();
                int wb = 0;
                foreach (Process p in ps)
                {
                    try { if (!string.IsNullOrEmpty(p.MainWindowTitle)) wb++; } catch { }
                }
                h.OpenWorkbooks = wb.ToString(); // 实为"有窗口的进程数"，见 Verify 的精确窗口统计
            }
            catch { h.ExcelProcesses = "?"; h.OpenWorkbooks = "?"; }

            // 评分：每个 ProgId 25 分（ddeexec/application/topic/command 各 6.25）
            double score = 0;
            if (total > 0) score = 100.0 * CountOk(h.Checks) / (total * 4);
            h.Score = (int)Math.Round(score);
            if (score >= 90) h.Verdict = "健康：DDE 合并机制完整（逐项值校验通过；安全加固项独立监测）";
            else if (score >= 50) h.Verdict = "部分受损：部分 ProgId 的 DDE 值不正确";
            else h.Verdict = "受损：多实例与复制粘贴异常的根因";
            return h;
        }

        static int CountOk(List<CheckItem> checks)
        {
            int n = 0;
            foreach (var c in checks) if (c.Ok) n++;
            return n;
        }

        // ── 报告 ────────────────────────────────────────────
        public static string ExportReport()
        {
            var sb = new StringBuilder();
            EnvInfo e = GetEnv();
            HealthInfo h = Diagnose();

            sb.AppendLine("Excel DDE 合并机制 - 诊断报告 v3");
            sb.AppendLine("生成时间: " + e.Now);
            sb.AppendLine("==================================================");
            sb.AppendLine("【环境指纹】");
            sb.AppendLine("  主机/用户     : " + e.Machine + " / " + e.User + "  管理员=" + e.Admin);
            sb.AppendLine("  操作系统      : " + e.Os + "  " + e.OsBits);
            sb.AppendLine("  .NET 运行时   : " + e.DotNet);
            sb.AppendLine("  Excel 路径    : " + e.ExcelPath);
            sb.AppendLine("  Excel 版本/位 : " + e.ExcelVersion + " / " + e.ExcelBits);
            sb.AppendLine("  Office 版本   : " + e.OfficeVersionsFound + " (主=" + h.OfficeVersion + ")");
            sb.AppendLine("  工具版本      : " + e.AppVersion);
            sb.AppendLine();
            sb.AppendLine("【健康评分】" + h.Score + "/100  -> " + h.Verdict);
            sb.AppendLine("  评分规则: 4 个 Excel ProgId × 4 项（ddeexec/application/topic/command）× 6.25 分");
            sb.AppendLine("  DDECleaned 是安全加固标记，不计入评分");
            sb.AppendLine("  Excel 进程: " + h.ExcelProcesses + " 个, 有窗口进程: " + h.OpenWorkbooks + " 个");
            sb.AppendLine();
            sb.AppendLine("【逐项精确校验】");
            foreach (CheckItem c in h.Checks)
            {
                sb.AppendLine("  " + (c.Ok ? "[OK]  " : "[FAIL]") + " " + c.ProgId + " \\ " + c.Item);
                if (!c.Ok)
                {
                    sb.AppendLine("         期望: " + c.Expected);
                    sb.AppendLine("         实际: " + c.Actual);
                }
            }
            sb.AppendLine();
            sb.AppendLine("【注册表原始快照】");
            foreach (RegValue v in CaptureCurrent())
            {
                sb.AppendLine("  " + v.Path + "  [" + v.Display + "]  kind=" + v.Kind);
                sb.AppendLine("      = " + v.Value);
            }
            sb.AppendLine();
            sb.AppendLine("【安全加固项监测】（不影响合并评分）");
            sb.AppendLine("  DDECleaned    = " + h.DdeCleaned + "  (kind=" + h.DdeCleanedKind + ")");
            sb.AppendLine("  DDEAllowed    = " + h.DdeAllowed + "  " + (h.DdeAllowed == "<MISSING>" ? "未设置=默认允许" : ""));
            sb.AppendLine("  DontUpdateLinks = " + h.DontUpdateLinks);
            sb.AppendLine("  IgnoreOtherApplications = " + h.IgnoreOtherApps);
            sb.AppendLine("  已启用安全加固项数量: " + h.HardeningCount);
            foreach (string n in h.SafetyNotes) sb.AppendLine("  注意: " + n);
            sb.AppendLine("==================================================");
            return sb.ToString();
        }
    }
}
