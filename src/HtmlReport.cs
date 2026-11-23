using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelDdeApp
{
    // ── 自包含 HTML 诊断报告：评分趋势 + 事件时间线 + 环境指纹 ──
    // 不依赖任何外部 JS/CSS，可直接双击打开或发给同事。
    internal static class HtmlReport
    {
        static string Esc(string s)
        {
            if (s == null) return "";
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    .Replace("\"", "&quot;").Replace("'", "&#39;");
        }

        class Point
        {
            public string Time;
            public int Score;
            public string Raw;
        }

        static List<Point> ReadTrend()
        {
            var list = new List<Point>();
            try
            {
                string dir = Log.Dir;
                if (!Directory.Exists(dir)) return list;
                string[] files = Directory.GetFiles(dir, "excel-dde-*.log");
                Array.Sort(files);
                Regex rx = new Regex(@"score=(\d{1,3})", RegexOptions.IgnoreCase);
                foreach (string f in files)
                {
                    foreach (string line in File.ReadAllLines(f, new UTF8Encoding(false)))
                    {
                        Match m = rx.Match(line);
                        if (!m.Success) continue;
                        int sc;
                        if (!int.TryParse(m.Groups[1].Value, out sc)) continue;
                        if (sc < 0 || sc > 100) continue;
                        string time = line.Length >= 23 ? line.Substring(0, 23) : "";
                        list.Add(new Point { Time = time, Score = sc, Raw = line });
                    }
                }
            }
            catch { }
            if (list.Count > 120) list.RemoveRange(0, list.Count - 120);
            return list;
        }

        static List<string> ReadTimeline()
        {
            var res = new List<string>();
            try
            {
                foreach (string l in Log.Tail(300))
                {
                    if (l.IndexOf("CLICK", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        l.IndexOf("DONE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        l.IndexOf("REPAIR", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        l.IndexOf("ROLLBACK", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        l.IndexOf("HARDEN", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        l.IndexOf("VERIFY", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        l.IndexOf("BACKUP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        l.IndexOf("ERROR", StringComparison.OrdinalIgnoreCase) >= 0)
                        res.Add(l);
                }
            }
            catch { }
            if (res.Count > 80) res.RemoveRange(0, res.Count - 80);
            res.Reverse();
            return res;
        }

        static string TrendSvg(List<Point> pts)
        {
            if (pts.Count == 0)
                return "<div class='empty'>日志中还没有 score 记录。点一次『刷新状态』或『一键修复』后即可看到趋势。</div>";

            int w = 940, h = 230, left = 46, right = 16, top = 18, bottom = 40;
            int pw = w - left - right, ph = h - top - bottom;
            var sb = new StringBuilder();
            sb.Append("<svg viewBox='0 0 " + w + " " + h + "' width='100%' role='img' aria-label='健康评分趋势'>");
            sb.Append("<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='#ffffff'/>");

            // 网格与 Y 轴刻度
            for (int v = 0; v <= 100; v += 20)
            {
                int y = top + ph - (int)(ph * v / 100.0);
                sb.Append("<line x1='" + left + "' y1='" + y + "' x2='" + (w - right) + "' y2='" + y +
                          "' stroke='#e8ecf1' stroke-width='1'/>");
                sb.Append("<text x='" + (left - 10) + "' y='" + (y + 4) + "' text-anchor='end' font-size='11' fill='#6b7280'>" + v + "</text>");
            }
            // 90 分健康线
            int y90 = top + ph - (int)(ph * 0.9);
            sb.Append("<line x1='" + left + "' y1='" + y90 + "' x2='" + (w - right) + "' y2='" + y90 +
                      "' stroke='#2e7d32' stroke-width='1' stroke-dasharray='5,4' opacity='0.65'/>");
            sb.Append("<text x='" + (w - right - 4) + "' y='" + (y90 - 6) + "' text-anchor='end' font-size='10' fill='#2e7d32'>健康线 90</text>");

            int n = pts.Count;
            var xy = new List<int[]>();
            for (int i = 0; i < n; i++)
            {
                int x = left + (n == 1 ? pw / 2 : (int)(pw * i / (double)(n - 1)));
                int y = top + ph - (int)(ph * pts[i].Score / 100.0);
                xy.Add(new int[] { x, y });
            }
            // 折线
            for (int i = 1; i < n; i++)
            {
                sb.Append("<line x1='" + xy[i - 1][0] + "' y1='" + xy[i - 1][1] + "' x2='" + xy[i][0] +
                          "' y2='" + xy[i][1] + "' stroke='#2563eb' stroke-width='2' stroke-linejoin='round'/>");
            }
            // 点 + 首尾分数标签
            for (int i = 0; i < n; i++)
            {
                int[] p = xy[i];
                string color = pts[i].Score >= 90 ? "#2e7d32" : (pts[i].Score >= 50 ? "#e67814" : "#c62828");
                sb.Append("<circle cx='" + p[0] + "' cy='" + p[1] + "' r='3.2' fill='" + color + "'/>");
                if (i == 0 || i == n - 1 || pts[i].Score != pts[Math.Max(0, i - 1)].Score)
                {
                    sb.Append("<text x='" + p[0] + "' y='" + (p[1] - 9) + "' text-anchor='middle' font-size='10' font-weight='bold' fill='" + color + "'>" + pts[i].Score + "</text>");
                }
            }
            sb.Append("<text x='" + left + "' y='" + (h - 12) + "' font-size='10' fill='#6b7280'>" + Esc(pts[0].Time) + "</text>");
            sb.Append("<text x='" + (w - right) + "' y='" + (h - 12) + "' text-anchor='end' font-size='10' fill='#6b7280'>" + Esc(pts[n - 1].Time) + "</text>");
            sb.Append("</svg>");
            return sb.ToString();
        }

        public static string Build()
        {
            var h = RegEngine.Diagnose();
            var e = RegEngine.GetEnv();
            var sb = new StringBuilder();

            sb.Append("<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>");
            sb.Append("<meta name='viewport' content='width=device-width, initial-scale=1'>");
            sb.Append("<title>Excel DDE 诊断报告 - " + Esc(e.Now) + "</title>");
            sb.Append("<style>");
            sb.Append("*{box-sizing:border-box}body{margin:0;background:#f5f7fa;color:#1f2937;font-family:'Microsoft YaHei','Segoe UI',sans-serif;line-height:1.6}");
            sb.Append(".wrap{max-width:1080px;margin:0 auto;padding:28px 22px 60px}");
            sb.Append("h1{font-size:25px;margin:0 0 4px;color:#111827}h2{font-size:16px;margin:26px 0 12px;padding-left:10px;border-left:4px solid #2563eb;color:#1f2937}");
            sb.Append(".sub{color:#6b7280;font-size:13px;margin-bottom:22px}");
            sb.Append(".card{background:#fff;border:1px solid #e5e9ef;border-radius:10px;padding:18px 20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(16,24,40,.04)}");
            sb.Append(".score{display:flex;align-items:center;gap:22px}.num{font-size:56px;font-weight:800;line-height:1}.verdict{font-size:15px;font-weight:700}.rule{font-size:12px;color:#6b7280;margin-top:6px}");
            sb.Append("table{width:100%;border-collapse:collapse;font-size:13px}th,td{text-align:left;padding:8px 10px;border-bottom:1px solid #eef1f5;vertical-align:top}th{background:#f8fafc;color:#475569;font-weight:700}");
            sb.Append(".ok{color:#2e7d32;font-weight:700}.fail{color:#c62828;font-weight:700}.muted{color:#9ca3af}.mono{font-family:Consolas,monospace;font-size:12px;word-break:break-all}");
            sb.Append(".tl{max-height:420px;overflow:auto;font-family:Consolas,monospace;font-size:11.5px;white-space:pre-wrap;background:#fbfcfe;border:1px solid #eef1f5;border-radius:6px;padding:12px}");
            sb.Append(".empty{color:#9ca3af;padding:26px;text-align:center;font-size:13px}");
            sb.Append(".foot{color:#9ca3af;font-size:11px;margin-top:28px;text-align:center}");
            sb.Append(".pill{display:inline-block;padding:2px 9px;border-radius:12px;font-size:11px;font-weight:700;margin-right:6px}");
            sb.Append("</style></head><body><div class='wrap'>");

            sb.Append("<h1>Excel DDE 合并机制 · 诊断报告</h1>");
            sb.Append("<div class='sub'>生成时间 " + Esc(e.Now) + " &nbsp;·&nbsp; 工具 v" + Esc(e.AppVersion) +
                      " &nbsp;·&nbsp; " + Esc(e.Machine) + "/" + Esc(e.User) + "</div>");

            // 评分卡
            string color = h.Score >= 90 ? "#2e7d32" : (h.Score >= 50 ? "#e67814" : "#c62828");
            sb.Append("<div class='card score'><div class='num' style='color:" + color + "'>" + h.Score + "</div><div>");
            sb.Append("<div class='verdict' style='color:" + color + "'>" + Esc(h.Verdict) + "</div>");
            sb.Append("<div class='rule'>评分规则：4 个 Excel ProgId × 4 项（ddeexec / application / topic / command）× 6.25 分。" +
                      "DDECleaned 是安全加固标记，不计入评分。Excel 进程 " + Esc(h.ExcelProcesses) +
                      " 个，有窗口进程 " + Esc(h.OpenWorkbooks) + " 个。</div>");
            sb.Append("</div></div>");

            // 趋势
            sb.Append("<h2>健康评分趋势（来自历史日志，最近 120 条）</h2><div class='card'>");
            sb.Append(TrendSvg(ReadTrend()));
            sb.Append("</div>");

            // 环境指纹
            sb.Append("<h2>环境指纹</h2><div class='card'><table>");
            sb.Append("<tr><th style='width:150px'>项目</th><th>值</th></tr>");
            string[,] env = new string[,]
            {
                { "操作系统", e.Os + " (" + e.OsBits + ")" },
                { ".NET 运行时", e.DotNet },
                { "Excel 路径", e.ExcelPath },
                { "Excel 版本 / 位数", e.ExcelVersion + " / " + e.ExcelBits },
                { "Office 版本（已安装）", e.OfficeVersionsFound + "，主版本 " + h.OfficeVersion },
                { "运行权限", e.Admin ? "管理员" : "普通用户" },
                { "工具版本", e.AppVersion }
            };
            for (int i = 0; i < env.GetLength(0); i++)
                sb.Append("<tr><td>" + Esc(env[i, 0]) + "</td><td class='mono'>" + Esc(env[i, 1]) + "</td></tr>");
            sb.Append("</table></div>");

            // 逐项校验
            sb.Append("<h2>逐项精确校验</h2><div class='card'><table>");
            sb.Append("<tr><th style='width:70px'>结果</th><th style='width:230px'>ProgId</th><th style='width:105px'>项目</th><th>期望 / 实际</th></tr>");
            foreach (CheckItem c in h.Checks)
            {
                sb.Append("<tr><td class='" + (c.Ok ? "ok" : "fail") + "'>" + (c.Ok ? "OK" : "FAIL") + "</td>");
                sb.Append("<td class='mono'>" + Esc(c.ProgId) + "</td><td>" + Esc(c.Item) + "</td><td>");
                if (c.Ok) sb.Append("<span class='muted'>" + Esc(c.Actual) + "</span>");
                else sb.Append("<span class='fail'>期望: " + Esc(c.Expected) + "</span><br><span class='mono'>实际: " + Esc(c.Actual) + "</span>");
                sb.Append("</td></tr>");
            }
            sb.Append("</table></div>");

            // 安全项
            sb.Append("<h2>安全加固监测（不计入合并评分）</h2><div class='card'><table>");
            sb.Append("<tr><th style='width:210px'>加固项</th><th style='width:90px'>当前值</th><th style='width:70px'>类型</th><th>说明</th></tr>");
            foreach (var t in DdeTargets.SafetyTargets)
            {
                string kind;
                string val = RegEngine.ReadValue(t.Path, t.Name, out kind);
                bool on = val != "<MISSING>" && val != "0";
                sb.Append("<tr><td class='mono'>" + Esc(t.Name) + "</td>");
                sb.Append("<td class='" + (on ? "ok" : "muted") + "'>" + Esc(val == "<MISSING>" ? "未设置" : val) + "</td>");
                sb.Append("<td class='muted'>" + Esc(kind == "Missing" ? "—" : kind) + "</td>");
                sb.Append("<td>" + Esc(t.Desc) + "</td></tr>");
            }
            sb.Append("</table><div class='rule' style='margin-top:10px'>已启用加固项 " + h.HardeningCount + " 个。" +
                      "这些设置用于防 DDE 攻击，实测不影响文件打开合并，因此不扣分。</div></div>");

            // 事件时间线
            sb.Append("<h2>关键事件时间线（最近 80 条）</h2><div class='card'><div class='tl'>");
            foreach (string l in ReadTimeline()) sb.Append(Esc(l) + "\n");
            sb.Append("</div></div>");

            sb.Append("<div class='foot'>本报告由 Excel DDE 合并修复工具自动生成 · 修改注册表前请务必备份 · 内容仅反映本机当前状态</div>");
            sb.Append("</div></body></html>");
            return sb.ToString();
        }
    }
}
