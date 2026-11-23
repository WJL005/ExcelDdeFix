using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ExcelDdeApp
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // 有参数 => 命令行模式（不弹界面，可用于自动化 / CI）
            if (args != null && args.Length > 0)
            {
                Environment.Exit(Cli.Run(args));
                return;
            }

            // 单实例保护：防止两个实例同时写注册表造成竞态。
            // 带重试是为了兼容『管理员重启』：UAC 新进程启动时旧实例可能尚未完全释放。
            bool created = false;
            Mutex single = null;
            for (int i = 0; i < 12; i++)
            {
                single = new Mutex(true, "ExcelDdeFix_SingleInstance_v3", out created);
                if (created) break;
                single.Close();
                single = null;
                Thread.Sleep(300);
            }
            if (!created)
            {
                MessageBox.Show("已有另一个 Excel DDE 修复工具实例正在运行，请先关闭它。",
                    "单实例保护", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try { RunGui(); }
            finally { if (single != null) { GC.KeepAlive(single); single.ReleaseMutex(); single.Close(); } }
        }

        static void RunGui()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += (s, e) =>
            {
                try { Log.Exception("UI线程未捕获", e.Exception); } catch { }
                MessageBox.Show(e.Exception.Message, "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                try { Log.Exception("进程级崩溃", e.ExceptionObject as Exception); } catch { }
            };
            Log.Cleanup(14);
            Log.WriteEx("INFO", "=== 程序启动 ===",
                "session=" + Log.SessionId +
                " machine=" + Environment.MachineName + " user=" + Environment.UserName +
                " os=" + Environment.OSVersion.VersionString + " exe=" + Application.ExecutablePath +
                " admin=" + MainForm.IsAdmin());
            Log.Write("INFO", "启动状态 | " + Log.StatusSnapshot());
            Application.Run(new MainForm());
            Log.Write("INFO", "=== 程序正常退出 ===");
        }
    }

    // ── 详细日志引擎 ─────────────────────────────────────────
    // 格式: 时间(毫秒) | 级别 | [当前操作] | 详情
    // 每次按钮点击自动生成 CLICK(开始) / DONE(结束+耗时) 两条，
    // 其间所有 Write 自动带上 [按钮名] 前缀；对话框选择、异常堆栈、前后状态快照均记录。
    static class Log
    {
        public static string Dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
        public static string TodayFile { get { return Path.Combine(Dir, "excel-dde-" + DateTime.Now.ToString("yyyyMMdd") + ".log"); } }

        static string _action = "";   // 当前操作标签，由按钮点击包装设置

        // 会话ID：每次启动唯一，便于从混合日志里筛出一次完整运行
        public static readonly string SessionId = Guid.NewGuid().ToString("N").Substring(0, 8);
        public static string JsonFile { get { return Path.Combine(Dir, "excel-dde-" + DateTime.Now.ToString("yyyyMMdd") + ".jsonl"); } }

        static string Stamp(string level)
        {
            string l = level ?? "INFO";
            if (l.Length > 9) l = l.Substring(0, 9);
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " + l.PadRight(9);
        }

        static string Prefix()
        {
            return string.IsNullOrEmpty(_action) ? "" : "[" + _action + "] ";
        }

        public static void SetAction(string action) { _action = action ?? ""; }
        public static void ClearAction() { _action = ""; }

        static string J(string s)
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

        // 结构化日志：每行一个 JSON 对象，方便 jq / Python / Excel PowerQuery 分析
        static void AppendJson(string ts, string level, string detail)
        {
            try
            {
                var sb = new StringBuilder();
                sb.Append("{\"ts\":").Append(J(ts));
                sb.Append(",\"sid\":").Append(J(SessionId));
                sb.Append(",\"level\":").Append(J(level));
                sb.Append(",\"action\":").Append(J(_action));
                sb.Append(",\"detail\":").Append(J(detail));
                sb.Append("}");
                File.AppendAllText(JsonFile, sb.ToString() + Environment.NewLine, new UTF8Encoding(false));
            }
            catch { }
        }

        // 常规一行日志（自动带当前操作前缀），同时写入 .log 与 .jsonl
        public static void Write(string level, string detail)
        {
            try
            {
                if (!Directory.Exists(Dir)) Directory.CreateDirectory(Dir);
                string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string l = level ?? "INFO";
                if (l.Length > 9) l = l.Substring(0, 9);
                string line = ts + " | " + l.PadRight(9) + " | " + Prefix() + detail;
                File.AppendAllText(TodayFile, line + Environment.NewLine, new UTF8Encoding(false));
                AppendJson(ts, l.Trim(), detail);
            }
            catch { }
        }

        // 带 key=value 明细的一行日志
        public static void WriteEx(string level, string detail, string fields)
        {
            Write(level, string.IsNullOrEmpty(fields) ? detail : detail + " | " + fields);
        }

        // 逐项 diff 日志：把注册表变化逐条写清楚，而不是只记一句"修复成功"
        public static void Diff(string label, List<RegValue> before, List<RegValue> after)
        {
            foreach (string d in RegEngine.Diff(before, after)) Write("DIFF", label + " " + d);
        }

        // 当前注册表/进程状态快照，点击前后各记一次，便于对比
        public static string StatusSnapshot()
        {
            try
            {
                var h = RegEngine.Diagnose();
                return string.Format("score={0} dde={1}/{2} cmdDde={3} DDECleaned={4}({5}) excelProc={6} win={7} hardening={8}",
                    h.Score, h.DdeOk, h.ExcelProgIds, h.CmdWithDde, h.DdeCleaned, h.DdeCleanedKind,
                    h.ExcelProcesses, h.OpenWorkbooks, h.HardeningCount);
            }
            catch (Exception ex) { return "status_error=" + ex.Message; }
        }

        // 按钮点击开始：打标签 + 记录点击行（含点击前状态）
        public static void BeginClick(string button, string before)
        {
            _action = button ?? "";
            Write("CLICK", ">>> 点击[" + (_action == "" ? "?" : _action) + "] | 前: " + before);
        }

        // 按钮点击结束：记录耗时与点击后状态，清除标签
        public static void EndClick(string button, bool ok, long ms, string after)
        {
            Write("DONE", "<<< [" + button + "] " + (ok ? "成功" : "失败") +
                " | 耗时=" + ms + "ms | 后: " + after);
            _action = "";
        }

        // 对话框选择（用户点了 OK / Cancel / Yes / No）
        public static void Choice(string question, string answer)
        {
            Write("CHOICE", "用户选择=" + answer + " | 问题: " + question);
        }

        // 异常全量记录：类型、消息、来源、调用堆栈、内部异常
        public static void Exception(string where, Exception ex)
        {
            // where 可能已等于当前操作名，避免前缀重复：此时用"操作内异常"占位
            string tag = (!string.IsNullOrEmpty(_action) && where == _action) ? "操作内异常" : "[" + where + "]";
            if (ex == null) { Write("ERROR", tag + " 异常对象为空"); return; }
            try
            {
                Write("ERROR", tag + " 异常: " + ex.GetType().FullName + ": " + ex.Message);
                Write("ERROR", "  Source=" + ex.Source + " TargetSite=" + (ex.TargetSite == null ? "?" : ex.TargetSite.ToString()));
                string st = ex.StackTrace;
                if (string.IsNullOrEmpty(st)) st = "(无堆栈)";
                foreach (string ln in st.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                    Write("ERROR", "  ST " + ln.Trim());
                if (ex.InnerException != null)
                    Write("ERROR", "  InnerException=" + ex.InnerException.GetType().FullName + ": " + ex.InnerException.Message);
            }
            catch { }
        }

        public static void Cleanup(int days)
        {
            try
            {
                if (!Directory.Exists(Dir)) return;
                DateTime cut = DateTime.Now.AddDays(-days);
                foreach (string f in Directory.GetFiles(Dir, "*.log"))
                    try { if (File.GetLastWriteTime(f) < cut) File.Delete(f); } catch { }
            }
            catch { }
        }
        public static string[] Tail(int n)
        {
            try
            {
                if (!File.Exists(TodayFile)) return new string[0];
                string[] all = File.ReadAllLines(TodayFile, new UTF8Encoding(false));
                var res = new List<string>();
                for (int i = all.Length - 1; i >= 0 && res.Count < n; i--) res.Add(all[i]);
                return res.ToArray();
            }
            catch { return new string[0]; }
        }
    }

    class MainForm : Form
    {
        // 控件
        Label lblScore, lblVerdict, lblDetail;
        DataGridView grid;
        DataGridView safetyGrid;      // 安全加固监测表（分析报告新增）
        ComboBox cmboBackup;
        CheckBox chkAutoRefresh;
        System.Windows.Forms.Timer timer;
        TextBox txtLog;
        Label lblHeader;
        Label lblSafetySummary;       // 安全监测摘要行

        static readonly Color BG = Color.FromArgb(248, 249, 251);
        static readonly Color CARD = Color.White;
        static readonly Color ACCENT = Color.FromArgb(37, 99, 235);
        static readonly Color GREEN = Color.FromArgb(46, 125, 50);
        static readonly Color RED = Color.FromArgb(198, 40, 40);
        static readonly Color AMBER = Color.FromArgb(230, 120, 20);
        static readonly Color LINE = Color.FromArgb(220, 224, 230);

        public MainForm()
        {
            Text = "Excel DDE 合并修复工具 v3 · " + Environment.MachineName + " · " + Environment.UserName;
            Font = new Font("Microsoft YaHei", 9F);
            BackColor = BG;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1120, 820);
            Size = new Size(1220, 940);

            BuildHeader();
            BuildStatusCard();
            BuildGrid();
            BuildSafetyPanel();
            BuildActionPanel();
            BuildBackupPanel();
            BuildLogPanel();
            BuildTimer();

            RefreshAll();
        }

        static Font F(float size, FontStyle style = FontStyle.Regular)
        {
            return new Font("Microsoft YaHei", size, style);
        }

        void BuildHeader()
        {
            var header = new Panel { Dock = DockStyle.Top, Height = 78, BackColor = CARD, Padding = new Padding(18, 8, 18, 8) };
            header.Controls.Add(new Label
            {
                Text = "Excel DDE 合并修复工具 v3 · 单实例 / 精确评分 / 实测验证",
                Font = F(19, FontStyle.Bold), ForeColor = Color.FromArgb(31, 35, 41),
                Dock = DockStyle.Top, Height = 34, TextAlign = ContentAlignment.MiddleLeft
            });
            lblHeader = new Label
            {
                Font = F(10), ForeColor = Color.FromArgb(107, 114, 128),
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            };
            header.Controls.Add(lblHeader);
            Controls.Add(header);
        }

        void BuildStatusCard()
        {
            var card = new Panel { Dock = DockStyle.Top, Height = 104, BackColor = CARD, Margin = new Padding(0), Padding = new Padding(16, 10, 16, 10) };

            lblScore = new Label { Font = F(40, FontStyle.Bold), Text = "--", Left = 18, Top = 14, Width = 130, Height = 56 };
            card.Controls.Add(lblScore);

            lblVerdict = new Label { Font = F(14, FontStyle.Bold), Text = "正在检测...", Left = 165, Top = 18, Width = 600, Height = 26 };
            card.Controls.Add(lblVerdict);

            lblDetail = new Label { Font = F(10), ForeColor = Color.FromArgb(90, 96, 106), Left = 165, Top = 48, Width = 900, Height = 44 };
            card.Controls.Add(lblDetail);

            Controls.Add(card);
        }

        void BuildGrid()
        {
            var gb = new GroupBox { Text = "  注册表实时状态（点刷新更新）  ", Dock = DockStyle.Top, Height = 250, Padding = new Padding(10), Font = F(10, FontStyle.Bold) };
            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = F(9)
            };
            grid.Columns.Add("colPath", "注册表位置");
            grid.Columns.Add("colName", "值名");
            grid.Columns.Add("colValue", "当前值");
            grid.Columns[0].Width = 300;
            grid.Columns[1].Width = 90;
            grid.Columns[2].Width = 500;
            grid.RowTemplate.Height = 22;
            gb.Controls.Add(grid);
            Controls.Add(gb);
        }

        // ── DDE 安全加固监测面板（分析报告新增）─────────────
        void BuildSafetyPanel()
        {
            var gb = new GroupBox
            {
                Text = "  DDE 安全加固监测（这些是安全设置，不是故障；独立监测、不计入合并评分）  ",
                Dock = DockStyle.Top, Height = 168, Padding = new Padding(10), Font = F(10, FontStyle.Bold)
            };

            lblSafetySummary = new Label
            {
                Left = 12, Top = 22, Width = 1100, Height = 20, Font = F(9, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 78, 121), Text = "安全加固项加载中…"
            };
            gb.Controls.Add(lblSafetySummary);

            safetyGrid = new DataGridView
            {
                Left = 12, Top = 44, Width = 1100, Height = 108,
                ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                BackgroundColor = Color.White, BorderStyle = BorderStyle.None,
                RowHeadersVisible = false, Font = F(9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            safetyGrid.Columns.Add("sName", "加固项");
            safetyGrid.Columns.Add("sValue", "当前值");
            safetyGrid.Columns.Add("sKind", "类型");
            safetyGrid.Columns.Add("sDesc", "说明");
            safetyGrid.Columns[0].Width = 200;
            safetyGrid.Columns[1].Width = 90;
            safetyGrid.Columns[2].Width = 70;
            safetyGrid.Columns[3].Width = 720;
            gb.Controls.Add(safetyGrid);

            Controls.Add(gb);
        }

        void BuildActionPanel()
        {
            var gb = new GroupBox { Text = "  修复与监控操作  ", Dock = DockStyle.Top, Height = 128, Padding = new Padding(12, 8, 12, 8), Font = F(10, FontStyle.Bold) };
            int x = 14, y = 22;

            MB(gb, "刷新状态", 100, x, y, RefreshAll, ACCENT); x += 110;
            MB(gb, "备份当前状态", 120, x, y, DoBackup, Color.FromArgb(21, 101, 192)); x += 130;
            MB(gb, "一键修复 DDE", 130, x, y, DoRepair, GREEN); x += 140;
            MB(gb, "回滚到备份", 110, x, y, DoRollback, RED); x += 120;

            MB(gb, "诊断报告", 96, x, y, DoExport, Color.FromArgb(97, 97, 97)); x += 106;
            MB(gb, "修正安全项类型", 130, x, y, DoFixHardeningType, Color.FromArgb(13, 148, 136)); x += 140;
            MB(gb, "Excel /regserver", 140, x, y, DoRegserver, Color.FromArgb(123, 31, 162)); x += 150;
            MB(gb, "管理员重启", 110, x, y, Elevate, Color.FromArgb(97, 97, 97));

            // 第二行：监控开关
            chkAutoRefresh = new CheckBox { Text = "自动刷新状态 (每 5 秒)", Left = 14, Top = 58, Width = 220, Font = F(9), ForeColor = Color.FromArgb(80, 86, 96) };
            gb.Controls.Add(chkAutoRefresh);

            var lblTip = new Label { Text = "提示：修复前会自动备份；任一步骤都可回滚。修改注册表请谨慎。", Left = 250, Top = 58, Width = 780, Font = F(9), ForeColor = AMBER };
            gb.Controls.Add(lblTip);

            // 第三行：闭环验证与开源配套
            int y2 = 88, x2 = 14;
            MB(gb, "实测合并验证", 124, x2, y2, DoVerify, Color.FromArgb(2, 132, 199)); x2 += 134;
            MB(gb, "导出HTML报告", 126, x2, y2, DoHtmlReport, Color.FromArgb(79, 70, 229)); x2 += 136;
            MB(gb, "版本与帮助", 108, x2, y2, DoAbout, Color.FromArgb(97, 97, 97)); x2 += 118;
            var lblVerify = new Label
            {
                Text = "实测验证 = 自动打开两个选定 xlsx → 数新增 Excel 进程 → 只关闭自己拉起的进程（不动你已开的表格）",
                Left = x2 + 6, Top = y2 + 6, Width = 760, Font = F(9), ForeColor = Color.FromArgb(107, 114, 128)
            };
            gb.Controls.Add(lblVerify);

            Controls.Add(gb);
        }

        void BuildBackupPanel()
        {
            var gb = new GroupBox { Text = "  备份管理  ", Dock = DockStyle.Top, Height = 74, Padding = new Padding(12, 8, 12, 8), Font = F(10, FontStyle.Bold) };
            cmboBackup = new ComboBox { Left = 14, Top = 24, Width = 620, DropDownStyle = ComboBoxStyle.DropDownList, Font = F(9) };
            gb.Controls.Add(cmboBackup);
            MB(gb, "刷新列表", 90, 650, 22, LoadBackups, Color.FromArgb(97, 97, 97));
            MB(gb, "查看此备份", 110, 750, 22, ViewBackup, Color.FromArgb(21, 101, 192));
            MB(gb, "打开备份目录", 116, 870, 22, () => StartSafe("explorer.exe", RegEngine.BackupDir), Color.FromArgb(97, 97, 97));
            MB(gb, "打开日志", 90, 996, 22, () => StartSafe("explorer.exe", Log.Dir), Color.FromArgb(97, 97, 97));
            Controls.Add(gb);
        }

        void BuildLogPanel()
        {
            var gb = new GroupBox { Text = "  操作日志（最新在上）  ", Dock = DockStyle.Fill, Padding = new Padding(10), Font = F(10, FontStyle.Bold) };
            txtLog = new TextBox
            {
                Multiline = true, ReadOnly = true, Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 9F),
                BackColor = Color.FromArgb(252, 252, 253), BorderStyle = BorderStyle.None
            };
            gb.Controls.Add(txtLog);
            Controls.Add(gb);
        }

        string _lastAutoSnap = "";

        void BuildTimer()
        {
            timer = new System.Windows.Forms.Timer { Interval = 5000 };
            timer.Tick += (s, e) =>
            {
                if (!chkAutoRefresh.Checked) return;
                RefreshStatus(); RefreshGrid(); LoadLog();
                string snap = Log.StatusSnapshot();
                if (snap != _lastAutoSnap)
                {
                    Log.Write("AUTO", "自动刷新检测到状态变化 | " + snap);
                    _lastAutoSnap = snap;
                }
            };
            chkAutoRefresh.CheckedChanged += (s, e) =>
            {
                timer.Enabled = chkAutoRefresh.Checked;
                Log.WriteEx("UI", "自动刷新开关=" + (chkAutoRefresh.Checked ? "开(5秒)" : "关"),
                    "action=" + (chkAutoRefresh.Checked ? "启动定时器" : "停止定时器"));
                if (chkAutoRefresh.Checked) _lastAutoSnap = Log.StatusSnapshot();
            };
        }

        Button MakeButton(string text, int width, int x, int y, Action onClick, Color bg)
        {
            var b = new Button
            {
                Text = text, Width = width, Height = 30, Left = x, Top = y,
                BackColor = bg, ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
                Font = F(9, FontStyle.Bold), Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            b.Click += (s, e) =>
            {
                var sw = Stopwatch.StartNew();
                Log.BeginClick(text, Log.StatusSnapshot());
                try
                {
                    onClick();
                    sw.Stop();
                    Log.EndClick(text, true, sw.ElapsedMilliseconds, Log.StatusSnapshot());
                }
                catch (Exception ex)
                {
                    sw.Stop();
                    Log.Exception(text, ex);
                    Log.EndClick(text, false, sw.ElapsedMilliseconds, Log.StatusSnapshot());
                    MessageBox.Show(ex.Message, "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    LoadLog();   // 每次点击后立即刷新日志面板
                }
            };
            return b;
        }

        // MakeButton 需要把按钮加入正确的容器：用一个重载
        Button MB(Control parent, string text, int width, int x, int y, Action onClick, Color bg)
        {
            var b = MakeButton(text, width, x, y, onClick, bg);
            parent.Controls.Add(b);
            return b;
        }

        // ── 刷新 ────────────────────────────────────────────
        void RefreshAll()
        {
            RefreshStatus();
            RefreshGrid();
            RefreshSafety();
            LoadBackups();
            LoadLog();
            lblHeader.Text = "运行时间 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                             "   |   Excel 路径: " + RegEngine.FindExcelExe() +
                             (IsAdmin() ? "   |   管理员权限" : "   |   普通权限");
            Log.Write("INFO", "刷新面板完成 | " + Log.StatusSnapshot() + " | hardening=" + RegEngine.Diagnose().HardeningCount);
        }

        void RefreshStatus()
        {
            var h = RegEngine.Diagnose();
            lblScore.Text = h.Score.ToString();
            lblScore.ForeColor = h.Score >= 90 ? GREEN : (h.Score >= 50 ? AMBER : RED);
            lblVerdict.Text = h.Verdict;
            lblVerdict.ForeColor = lblScore.ForeColor;
            int okCount = 0;
            foreach (var c in h.Checks) if (c.Ok) okCount++;
            lblDetail.Text = string.Format(
                "逐项值校验通过: {0}/{1} 项（4 ProgId × 4 项 × 6.25 分）   |   ddeexec 值正确: {2}/{1}   |   command /dde: {3}/{1}\n" +
                "Excel 进程: {4} 个   |   Office {5}   |   Excel: {6}   |   DDECleaned={7}({8}) 为安全加固标记，不计分",
                okCount, h.ExcelProgIds * 4, h.DdeOk, h.CmdWithDde, h.ExcelProcesses, h.OfficeVersion,
                string.IsNullOrEmpty(h.ExcelExe) ? "未找到" : Path.GetFileName(h.ExcelExe), h.DdeCleaned, h.DdeCleanedKind);
        }

        // 刷新安全加固监测表
        void RefreshSafety()
        {
            safetyGrid.Rows.Clear();
            var h = RegEngine.Diagnose();
            foreach (var t in DdeTargets.SafetyTargets)
            {
                string kind;
                string val = RegEngine.ReadValue(t.Path, t.Name, out kind);
                string shortName = t.Path.StartsWith("HKEY_CURRENT_USER\\") ? t.Path.Substring(18)
                                 : (t.Path.StartsWith("HKEY_LOCAL_MACHINE\\") ? "HKLM\\" + t.Path.Substring(19) : t.Path);
                int r = safetyGrid.Rows.Add(t.Name, val == "<MISSING>" ? "（未设置）" : val,
                                            kind == "Missing" ? "—" : kind, shortName + " — " + t.Desc);
                // 已加固=绿；DDEAllowed=0=琥珀（提示）；缺失=灰
                if (val == "<MISSING>" || val == "0")
                    safetyGrid.Rows[r].Cells[1].Style.ForeColor = val == "0" && t.Name == DdeTargets.S_DDEAllowed ? AMBER : Color.Gray;
                else
                    safetyGrid.Rows[r].Cells[1].Style.ForeColor = GREEN;
            }
            lblSafetySummary.Text = string.Format(
                "安全加固项：DDECleaned={0}({1})   DDEAllowed={2}   DontUpdateLinks={3}   IgnoreOtherApps={4}   |   已启用加固项 {5} 个   |   结论：这些设置防 DDE 攻击，实测不影响文件打开合并",
                h.DdeCleaned, h.DdeCleanedKind, Disp(h.DdeAllowed), Disp(h.DontUpdateLinks), Disp(h.IgnoreOtherApps), h.HardeningCount);
            foreach (string note in h.SafetyNotes)
                lblSafetySummary.Text += "   ⚠ " + note;
        }

        static string Disp(string v) { return v == "<MISSING>" ? "未设置" : v; }

        void RefreshGrid()
        {
            grid.Rows.Clear();
            foreach (RegValue v in RegEngine.CaptureCurrent())
            {
                int r = grid.Rows.Add(v.Path, v.Display, v.Value);
                if (v.Value == "<MISSING>")
                {
                    grid.Rows[r].Cells[2].Style.ForeColor = RED;
                    grid.Rows[r].Cells[2].Value = "【缺失】<MISSING>";
                }
                else if (v.Value.IndexOf("/dde", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         v.Value.IndexOf("[open", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    grid.Rows[r].Cells[2].Style.ForeColor = GREEN;
                }
            }
        }

        void LoadBackups()
        {
            cmboBackup.Items.Clear();
            var list = RegEngine.ListBackups();
            foreach (var b in list)
                cmboBackup.Items.Add(b.At.ToString("yyyy-MM-dd HH:mm:ss") + (string.IsNullOrEmpty(b.Label) ? "" : "  · " + b.Label));
            if (cmboBackup.Items.Count > 0) cmboBackup.SelectedIndex = 0;
        }

        void LoadLog()
        {
            txtLog.Text = string.Join(Environment.NewLine, Log.Tail(80));
            txtLog.SelectionStart = 0;
            txtLog.ScrollToCaret();
        }

        // ── 操作 ────────────────────────────────────────────
        void DoBackup()
        {
            var b = RegEngine.SaveBackup("修复前备份");
            int missing = 0;
            foreach (var v in b.Values) if (v.Kind == "Missing") missing++;
            Log.WriteEx("BACKUP", "手动备份完成",
                "file=" + b.FilePath + " items=" + b.Values.Count + " missing=" + missing +
                " bytes=" + (File.Exists(b.FilePath) ? new FileInfo(b.FilePath).Length.ToString() : "?"));
            Log.Write("BACKUP", "备份明细: " + BackupBrief(b));
            LoadBackups();
            MessageBox.Show("已备份 " + b.Values.Count + " 项注册表状态：\n" + b.FilePath,
                "备份完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 生成备份内容的简短摘要（每项 PROGID\位置=值），便于在日志里直接看到备份了什么
        string BackupBrief(DdeBackup b)
        {
            if (b == null) return "(空)";
            var sb = new StringBuilder();
            foreach (var v in b.Values)
            {
                string shortPath = v.Path;
                if (shortPath.StartsWith("HKEY_CLASSES_ROOT\\")) shortPath = shortPath.Substring(18);
                else if (shortPath.StartsWith("HKEY_CURRENT_USER\\")) shortPath = shortPath.Substring(18);
                string val = v.Value.Length > 40 ? v.Value.Substring(0, 40) + "…" : v.Value;
                sb.Append(shortPath).Append("[").Append(v.Display).Append("]=").Append(val).Append("; ");
            }
            return sb.ToString();
        }

        void DoRepair()
        {
            int score = RegEngine.Diagnose().Score;
            Log.WriteEx("REPAIR", "修复前检查 | " + Log.StatusSnapshot(), "score=" + score);

            if (score >= 90)
            {
                var r = MessageBox.Show("当前状态看起来已健康（评分 " + score + "），仍要修复吗？",
                    "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                Log.Choice("状态已健康(score=" + score + ")仍要修复？", r.ToString());
                if (r == DialogResult.No) { Log.Write("INFO", "用户放弃修复（健康确认选否）"); return; }
            }

            var c = MessageBox.Show(
                    "将执行以下修复（v2，已修正安全项处理）：\n\n" +
                    "1. 为 .xlsx/.xls/.xlsm/.xlsb 恢复 ddeexec 合并注册\n" +
                    "2. 把打开命令改为 \"EXCEL.EXE\" /dde \"%1\"\n" +
                    "3. 不再把 DDECleaned 清零（它是安全加固标记，不影响合并）\n" +
                    "   仅纠正被旧版写坏的类型：String → DWord 1\n\n" +
                    "修复前会自动备份，可随时回滚。是否继续？",
                    "确认修复 DDE", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            Log.Choice("确认执行 DDE 修复", c.ToString());
            if (c != DialogResult.OK) { Log.Write("INFO", "用户取消修复（确认框非 OK）"); return; }

            // 自动备份
            var b = RegEngine.SaveBackup("修复前自动备份");
            Log.WriteEx("BACKUP", "修复前自动备份完成",
                "file=" + b.FilePath + " items=" + b.Values.Count + " missing=" + CountMissing(b));

            string exe = RegEngine.FindExcelExe(true);
            if (exe == null)
            {
                Log.Write("ERROR", "修复中止：未能定位 EXCEL.EXE（避免写入错误路径）");
                MessageBox.Show("未能定位 EXCEL.EXE（App Paths / COM / 已知路径都找不到），已中止修复，避免写入错误路径。",
                    "修复中止", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Log.WriteEx("REPAIR", "开始写入 DDE 注册", "exe=" + exe + " mode=dde bits=" + RegEngine.GetPeArchitecture(exe));
            var errs = RegEngine.Repair("dde", exe);
            var afterVals = RegEngine.CaptureCurrent();
            Log.Diff("修复变更", b.Values, afterVals);
            int scoreAfter = RegEngine.Diagnose().Score;
            Log.WriteEx("REPAIR", "DDE 修复结果",
                "score=" + score + "->" + scoreAfter + " errs=" + errs.Count +
                " exe=" + exe + " detail=" + (errs.Count == 0 ? "全部成功" : string.Join(";", errs.ToArray())));

            RefreshAll();
            MessageBox.Show(
                errs.Count == 0
                    ? "修复完成！评分 " + score + " → " + scoreAfter + "。\n请关闭所有 Excel 后重新打开两个 xlsx 测试。\n若仍异常，可点『回滚到备份』还原。"
                    : "修复完成但有 " + errs.Count + " 项失败：\n" + string.Join("\n", errs.ToArray()),
                "修复结果", MessageBoxButtons.OK, errs.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        int CountMissing(DdeBackup b)
        {
            int n = 0;
            if (b != null) foreach (var v in b.Values) if (v.Kind == "Missing") n++;
            return n;
        }

        // 修正被旧版工具写坏的 DDECleaned 类型：String -> 标准 DWord 1
        void DoFixHardeningType()
        {
            string kind;
            string val = RegEngine.ReadValue(DdeTargets.ExcelOptions(DdeTargets.PrimaryOfficeVersion()), DdeTargets.S_DDECleaned, out kind);
            Log.WriteEx("HARDEN", "修正安全项类型检查", "DDECleaned=" + val + " kind=" + kind);
            if (val == "<MISSING>")
            {
                Log.Write("INFO", "DDECleaned 不存在，无需修正");
                MessageBox.Show("DDECleaned 未设置（默认状态），无需修正。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (kind == "DWord")
            {
                Log.Write("INFO", "DDECleaned 已是标准 DWord，无需修正");
                MessageBox.Show("DDECleaned 已是标准 DWord 类型（值=" + val + "），无需修正。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var rc = MessageBox.Show(
                "检测到 DDECleaned 被写成了 " + kind + " 类型（当前值=" + val + "），\n" +
                "标准应为 DWord。旧版工具曾错误地用字符串写入，导致重启后被 Office 改写。\n\n" +
                "将修正为标准 DWord 1（保持「已加固」语义，不影响 DDE 合并）。是否继续？",
                "修正安全项类型", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            Log.Choice("确认修正 DDECleaned 类型 String->DWord", rc.ToString());
            if (rc != DialogResult.OK) { Log.Write("INFO", "用户取消修正安全项类型"); return; }

            var b = RegEngine.SaveBackup("修正安全项类型前备份");
            Log.WriteEx("BACKUP", "修正前自动备份", "file=" + b.FilePath + " items=" + b.Values.Count);
            var errs = RegEngine.Repair("clean", "");
            RegEngine.Diagnose();
            string newKind; string newVal = RegEngine.ReadValue(DdeTargets.ExcelOptions(DdeTargets.PrimaryOfficeVersion()), DdeTargets.S_DDECleaned, out newKind);
            Log.WriteEx("HARDEN", "安全项类型修正结果",
                "errs=" + errs.Count + " DDECleaned=" + newVal + " kind=" + newKind + " | " + Log.StatusSnapshot());
            RefreshAll();
            MessageBox.Show(errs.Count == 0
                    ? "已修正：DDECleaned = " + newVal + "，类型 = " + newKind + "。"
                    : "修正有 " + errs.Count + " 项失败：" + string.Join("\n", errs.ToArray()),
                "修正结果", MessageBoxButtons.OK, errs.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        void DoRollback()
        {
            DdeBackup b = null;
            if (cmboBackup.SelectedIndex >= 0)
            {
                var list = RegEngine.ListBackups();
                if (cmboBackup.SelectedIndex < list.Count) b = list[cmboBackup.SelectedIndex];
            }
            if (b == null)
            {
                Log.Write("WARN", "回滚中止：下拉框无可用备份 (index=" + cmboBackup.SelectedIndex + ", 共" + cmboBackup.Items.Count + "项)");
                MessageBox.Show("没有可回滚的备份，请先点『备份当前状态』。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return;
            }

            Log.WriteEx("ROLLBACK", "选中备份 | " + Log.StatusSnapshot(),
                "at=" + b.At.ToString("yyyy-MM-dd HH:mm:ss") + " file=" + b.FilePath +
                " items=" + b.Values.Count + " label=" + b.Label);

            var rc = MessageBox.Show(
                    "将把注册表还原到备份时间：\n" + b.At.ToString("yyyy-MM-dd HH:mm:ss") + "\n\n" +
                    "这会撤销该备份之后的所有修改（含 DDE 修复）。确认继续？",
                    "确认回滚", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            Log.Choice("确认回滚到 " + b.At.ToString("yyyy-MM-dd HH:mm:ss"), rc.ToString());
            if (rc != DialogResult.OK) { Log.Write("INFO", "用户取消回滚"); return; }

            // 回滚前再留一份保护快照：即使回滚本身出问题也能再退一步
            var guard = RegEngine.SaveBackup("回滚前保护快照");
            Log.WriteEx("BACKUP", "回滚前保护快照", "file=" + guard.FilePath + " items=" + guard.Values.Count);
            Log.Write("ROLLBACK", "备份完整性=" + b.Integrity);
            int before = RegEngine.Diagnose().Score;
            var errs = RegEngine.Restore(b);
            Log.Diff("回滚变更", b.Values, RegEngine.CaptureCurrent());
            int after = RegEngine.Diagnose().Score;
            Log.WriteEx("ROLLBACK", "回滚结果",
                "at=" + b.At.ToString("yyyy-MM-dd HH:mm:ss") + " items=" + b.Values.Count +
                " score=" + before + "->" + after + " errs=" + errs.Count +
                " detail=" + (errs.Count == 0 ? "全部成功" : string.Join(";", errs.ToArray())));
            RefreshAll();
            MessageBox.Show(errs.Count == 0 ? "已还原到备份状态（评分 " + before + " → " + after + "）。" : "还原部分失败：" + string.Join("\n", errs.ToArray()),
                "回滚结果", MessageBoxButtons.OK, errs.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        void ViewBackup()
        {
            DdeBackup b = null;
            if (cmboBackup.SelectedIndex >= 0)
            {
                var list = RegEngine.ListBackups();
                if (cmboBackup.SelectedIndex < list.Count) b = list[cmboBackup.SelectedIndex];
            }
            if (b == null) { Log.Write("WARN", "查看备份中止：未选择备份"); MessageBox.Show("请先选择一个备份。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            Log.WriteEx("VIEW", "查看备份内容", "at=" + b.At.ToString("yyyy-MM-dd HH:mm:ss") + " file=" + b.FilePath + " items=" + b.Values.Count);
            var sb = new StringBuilder();
            sb.AppendLine("备份时间: " + b.At.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine("项目数: " + b.Values.Count);
            sb.AppendLine("文件: " + b.FilePath);
            sb.AppendLine("------------------------------");
            foreach (var v in b.Values) sb.AppendLine(v.Path + " [" + v.Display + "] = " + v.Value);
            var f = new Form { Text = "备份内容", Size = new Size(900, 600), StartPosition = FormStartPosition.CenterParent, Font = F(9) };
            var t = new TextBox { Multiline = true, ReadOnly = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Both, Text = sb.ToString(), Font = new Font("Consolas", 9F) };
            f.Controls.Add(t);
            f.ShowDialog(this);
        }

        void DoExport()
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "diagnose-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt");
                File.WriteAllText(path, RegEngine.ExportReport(), new UTF8Encoding(false));
                Log.WriteEx("EXPORT", "诊断报告已导出",
                    "file=" + path + " bytes=" + new FileInfo(path).Length + " | " + Log.StatusSnapshot());
                MessageBox.Show("诊断报告已导出：\n" + path, "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { Log.Exception("导出诊断报告", ex); MessageBox.Show("导出失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        // ── 闭环验证：真正打开两个 xlsx，看是否合并为单进程 ──
        void DoVerify()
        {
            var ofd = new OpenFileDialog
            {
                Title = "选择两个 xlsx 做实测（验证 DDE 合并）",
                Filter = "Excel 工作簿|*.xlsx;*.xls;*.xlsm;*.xlsb",
                Multiselect = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (ofd.ShowDialog(this) != DialogResult.OK) { Log.Write("INFO", "实测验证取消：未选择文件"); return; }
            string[] files = ofd.FileNames;
            if (files.Length < 2) { MessageBox.Show("请至少选择两个文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            Log.WriteEx("VERIFY", "开始实测合并验证",
                "f1=" + files[0] + " f2=" + files[1] + " | " + Log.StatusSnapshot());
            var r = MergeVerifier.Run(files[0], files[1], 2500,
                delegate(string p) { Log.Write("VERIFY", p); });
            Log.WriteEx("VERIFY", "实测结论",
                "ok=" + r.Ok + " newProcs=" + r.NewExcelProcesses + " newWindows=" + r.NewExcelWindows +
                " pids=" + string.Join(",", r.NewPids.ConvertAll(x => x.ToString()).ToArray()) + " | " + r.Message);
            RefreshAll();
            MessageBox.Show(r.Message + "\n\n本次新拉起的进程: " +
                            (r.NewPids.Count == 0 ? "无" : string.Join(", ", r.NewPids.ConvertAll(x => x.ToString()).ToArray())) +
                            "\n（这些进程已被工具关闭，你原先打开的 Excel 不受影响）",
                r.Ok ? "实测通过：合并正常" : "实测未通过：未合并",
                MessageBoxButtons.OK, r.Ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        // ── 导出自包含 HTML 可视化报告 ──
        void DoHtmlReport()
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "diagnose-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".html");
                File.WriteAllText(path, HtmlReport.Build(), new UTF8Encoding(false));
                Log.WriteEx("EXPORT", "HTML 报告已导出",
                    "file=" + path + " bytes=" + new FileInfo(path).Length + " | " + Log.StatusSnapshot());
                StartSafe(path, "");
                MessageBox.Show("HTML 报告已导出并在浏览器中打开：\n" + path, "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { Log.Exception("导出 HTML 报告", ex); MessageBox.Show("导出失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        void DoAbout()
        {
            var e = RegEngine.GetEnv();
            string s = "Excel DDE 合并修复工具 v" + e.AppVersion + "\n\n" +
                "解决：打开多个 xlsx 时产生多个 EXCEL.EXE 进程、跨进程复制粘贴丢格式。\n\n" +
                "命令行（双击无参数进界面，带参数走命令行）：\n" +
                "  ExcelDdeFix.exe --diagnose [--json]\n" +
                "  ExcelDdeFix.exe --repair --dry-run\n" +
                "  ExcelDdeFix.exe --repair --yes     (需管理员)\n" +
                "  ExcelDdeFix.exe --verify 文件1.xlsx 文件2.xlsx\n" +
                "  ExcelDdeFix.exe --export --html\n" +
                "  ExcelDdeFix.exe --help\n\n" +
                "退出码: 0=成功 1=警告 2=参数错 3=权限不足 4=异常\n\n" +
                "环境: " + e.Os + " / Excel " + e.ExcelVersion + " (" + e.ExcelBits + ")\n" +
                "会话: " + Log.SessionId;
            Log.Write("INFO", "查看版本与帮助");
            MessageBox.Show(s, "版本与帮助", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void DoRegserver()
        {
            Log.WriteEx("REGSVR", "/regserver 权限检查", "admin=" + IsAdmin() + " | " + Log.StatusSnapshot());
            if (!IsAdmin())
            {
                Log.Write("WARN", "/regserver 中止：非管理员权限");
                MessageBox.Show("请先用『管理员重启』以管理员身份运行，再执行 /regserver。", "需要管理员", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var rc = MessageBox.Show("⚠ 请先保存并关闭所有 Excel 中未保存的工作！\n\n" +
                    "本操作会关闭当前所有 Excel 进程，并执行 EXCEL.EXE /unregserver 和 /regserver 重新注册 COM/DDE。\n" +
                    "工具会先尝试正常关闭窗口，只有未响应时才强制结束。\n\n继续执行？",
                    "确认 /regserver", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            Log.Choice("确认执行 /unregserver + /regserver", rc.ToString());
            if (rc != DialogResult.OK) { Log.Write("INFO", "用户取消 /regserver"); return; }
            try
            {
                string exe = RegEngine.FindExcelExe();
                var excelPs = Process.GetProcessesByName("EXCEL");
                string pids = "";
                foreach (Process p in excelPs)
                {
                    pids += p.Id + " ";
                    try
                    {
                        bool closed = p.CloseMainWindow();
                        Log.Write("REGSVR", "请求正常关闭 Excel PID=" + p.Id + " accepted=" + closed);
                    }
                    catch (Exception kex) { Log.Write("WARN", "请求关闭 Excel PID=" + p.Id + " 失败: " + kex.Message); }
                }
                System.Threading.Thread.Sleep(1800);
                foreach (Process p in excelPs)
                {
                    try
                    {
                        p.Refresh();
                        if (!p.HasExited)
                        {
                            p.Kill();
                            Log.Write("WARN", "Excel PID=" + p.Id + " 未响应正常关闭，已强制结束");
                        }
                    }
                    catch { }
                }
                Log.WriteEx("REGSVR", "已结束 Excel 进程", "count=" + excelPs.Length + " pids=" + pids.Trim());
                int code1 = RunWait(exe, "/unregserver");
                System.Threading.Thread.Sleep(500);
                int code2 = RunWait(exe, "/regserver");
                Log.WriteEx("REGSVR", "执行 /unregserver + /regserver 完成",
                    "exe=" + exe + " unregserver_exit=" + code1 + " regserver_exit=" + code2 + " | " + Log.StatusSnapshot());
                RefreshAll();
                MessageBox.Show("/regserver 执行完成（退出码 " + code1 + "/" + code2 + "），请重新打开 xlsx 测试。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { Log.Exception("/regserver", ex); MessageBox.Show("失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        int RunWait(string exe, string args)
        {
            try
            {
                var p = Process.Start(new ProcessStartInfo(exe, args) { UseShellExecute = false, CreateNoWindow = true });
                if (p != null)
                {
                    p.WaitForExit(15000);
                    int code = p.HasExited ? p.ExitCode : -1;
                    Log.WriteEx("REGSVR", "外部命令结束", "cmd=\"" + args + "\" exit=" + code + " timeout=" + (p.HasExited ? "否" : "是"));
                    return code;
                }
                Log.Write("WARN", "外部命令未启动: " + args);
                return -99;
            }
            catch (Exception ex) { Log.Exception("RunWait " + args, ex); return -98; }
        }

        void Elevate()
        {
            bool admin = IsAdmin();
            Log.WriteEx("ADMIN", "管理员重启请求", "currentAdmin=" + admin + " exe=" + Application.ExecutablePath);
            if (admin) { Log.Write("INFO", "已是管理员，无需重启"); MessageBox.Show("当前已是管理员权限。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            try
            {
                var psi = new ProcessStartInfo(Application.ExecutablePath)
                {
                    UseShellExecute = true, Verb = "runas"
                };
                Process.Start(psi);
                Log.Write("INFO", "已拉起 UAC 提权请求，当前实例即将退出");
                Application.Exit();
            }
            catch (Exception ex) { Log.Write("WARN", "UAC 提权取消或失败: " + ex.GetType().Name + ": " + ex.Message); }
        }

        internal static bool IsAdmin()
        {
            try { return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator); }
            catch { return false; }
        }

        static void StartSafe(string exe, string args)
        {
            Log.WriteEx("OPEN", "打开外部程序", "exe=" + exe + " args=" + args);
            try { Process.Start(new ProcessStartInfo(exe, args) { UseShellExecute = true }); }
            catch (Exception ex) { Log.Exception("打开外部程序 " + exe, ex); MessageBox.Show("打开失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
    }
}
