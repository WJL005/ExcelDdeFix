# Excel DDE 合并修复工具 · ExcelDdeFix

> 打开多个 `.xlsx` 时总蹦出好几个 `EXCEL.EXE`？在两个表格间复制粘贴，公式和格式总是丢？这个工具帮你查清楚、修回去、还能一眼验证。

一个面向 Windows 的诊断 / 修复工具，解决 **打开多个 xlsx 产生多个 `EXCEL.EXE` 进程、跨工作簿复制粘贴丢格式** 的问题。提供图形界面和命令行两种用法，所有操作**可备份、可回滚、可审计、可验证**。

![平台](https://img.shields.io/badge/platform-Windows-blue) ![.NET](https://img.shields.io/badge/.NET-Framework%204.0-5C2D91) ![许可证](https://img.shields.io/badge/license-MIT-green) ![版本](https://img.shields.io/badge/version-3.0.0-orange)

---

## 📑 目录

- [痛点：这是什么问题](#1-痛点这是什么问题)
- [功能一览](#2-功能一览)
- [看看它长什么样](#3-看看它长什么样示例产物)
- [快速开始](#4-快速开始)
- [评分规则](#5-评分规则)
- [命令行参考](#6-命令行参考)
- [日志与排障](#7-日志与排障)
- [编译](#8-编译)
- [安全说明](#9-安全说明)
- [已知限制](#10-已知限制)
- [参与贡献](#11-参与贡献)
- [许可证](#12-许可证)

---

## 1. 痛点：这是什么问题

Windows 依赖注册表中的 **DDE 注册**把"打开 xlsx 的请求"转发给已在运行的 Excel 实例。正常情况下，无论从资源管理器、浏览器下载目录，还是文件搜索工具打开，理论上都只有 **一个** `EXCEL.EXE` 进程，多个工作簿只是它里面的多个窗口。

当 `Excel.Sheet.12` 等 ProgId 的 `shell\open\ddeexec` 注册缺失，或 `command` 没有 `/dde` 参数时，这条转发通道断了：**每打开一个文件就新起一个 Excel 进程**。后果是：

| 场景 | 单进程（正常） | 多进程（异常） |
|---|---|---|
| 复制单元格 | 走 Excel 内部剪贴板，保留公式、格式、列宽、合并单元格、条件格式 | 退回 Windows 系统剪贴板，只剩纯文本 / HTML 表格，格式丢失 |
| 窗口位置 / 功能区 | 共享同一实例状态 | 各实例独立，看起来"像两套 Office" |
| 跨工作簿引用 | 正常 | 引用易错乱 |

**常见诱因**：安全加固脚本（HardeningKitty、DDE 攻击防护脚本）清理了 DDE 注册；优化软件改写了文件关联；不同入口（文件搜索工具、浏览器下载器、IM 软件）各自拉起进程。

### 关于 `DDECleaned` 的一个常见误区

排查中 `DDECleaned=1` 常与问题同时出现，容易被误当成"元凶"。实际上它是**安全加固标记**（DWord `1` = 已做 DDE 加固），**不影响文件打开合并**。本工具 v3 明确把它从评分中剔除，保留其安全语义，只纠正被旧工具写错的数据类型（String → DWord）。

> **一句话结论：DDE 合并是否健康，只看 `ddeexec` 与 `command /dde` 的实际值，不看安全加固标记。**

---

## 2. 功能一览

| 能力 | 说明 |
|---|---|
| 🎯 **精确健康评分** | 4 个 Excel ProgId × 4 项（`ddeexec` / `application` / `topic` / `command`）× 6.25 分，校验**值内容**而不只是键是否存在——值被写错也会判 FAIL |
| 🔍 **智能 Excel 探测** | App Paths 注册表 → COM `LocalServer32` → 多版本多路径（Office 2010/2013/2016+/C2R，32/64 位）→ 存在性校验；找不到就报错中止，**绝不写入猜测路径** |
| 💾 **备份 / 回滚 / 校验** | 每个备份带 SHA256 摘要和当时评分；回滚前校验哈希、自动再留保护快照；拒绝危险的整键删除 |
| ✅ **闭环实测验证** | 自动打开你选的两个 xlsx，统计**新增**的 Excel 进程数判断是否合并；只关闭工具自己拉起的进程，不碰你已打开的表格 |
| 📝 **全量可追溯日志** | 文本日志 + JSONL 结构化日志，带会话 ID、按钮点击前后状态快照、耗时、对话框选择、异常堆栈、注册表逐项 diff |
| 📊 **自包含 HTML 报告** | 评分趋势折线图、事件时间线、环境指纹、逐项校验结果，单文件可直接分享 |
| 🖥️ **GUI + CLI** | 图形界面适合手工排查，命令行适合自动化与 CI |
| 🔒 **单实例保护** | 命名 Mutex 防止两个实例并发写注册表造成竞态 |

---

## 3. 看看它长什么样（示例产物）

`examples/` 目录里是从真实运行**脱敏**后的示例，让你在下载前就能了解输出效果（所有机器名、用户名、私有路径、文件名均已泛化）：

| 示例文件 | 展示内容 |
|---|---|
| [`示例诊断报告.html`](examples/示例诊断报告.html) | 自包含 HTML 报告：健康评分、评分趋势折线图、逐项精确校验表、安全加固监测表、事件时间线。直接用浏览器打开 |
| [`示例操作日志.txt`](examples/示例操作日志.txt) | 人类可读日志：按钮点击 CLICK→DONE 链路、备份明细、耗时、状态快照 |
| [`示例日志片段.txt`](examples/示例日志片段.txt) | JSONL 结构化日志：每行一个 JSON，含 `ts` / `sid` / `level` / `action` / `detail` |

> 想看真实效果？下载 exe 在你自己的机器上跑一次 `--export --html`，生成的报告 Structure 与示例一致，只是填入你的真实环境数据。

---

## 4. 快速开始

### 4.1 图形界面（推荐大多数人）

双击 `ExcelDdeFix.exe`（建议先点右上角「管理员重启」获取管理员权限，修复 / 回滚需要）。

**标准流程：**

1. 点 **备份当前状态** —— 先留条后路
2. 点 **一键修复 DDE**（写前会自动再备份一次）
3. 关闭所有 Excel，重新打开两个 xlsx
4. 点 **实测合并验证**，选这两个文件，看是否提示"合并成功"
5. 效果不理想？选中刚才的备份，点 **回滚到备份**

界面分为六块：健康评分卡 / 注册表实时状态 / DDE 安全加固监测 / 操作按钮 / 备份管理 / 操作日志。勾选"自动刷新"后每 5 秒更新一次状态。

### 4.2 命令行（适合脚本 / CI）

**有参数走命令行，无参数进界面。** 退出码：`0=成功 1=警告/部分失败 2=参数错误 3=权限不足 4=异常`。

```bat
ExcelDdeFix.exe --diagnose                        :: 文本诊断报告
ExcelDdeFix.exe --diagnose --json                 :: JSON，方便脚本解析
ExcelDdeFix.exe --backup --label "改之前"          :: 备份
ExcelDdeFix.exe --repair --dry-run                :: 演练：只列出将写入什么，不改注册表
ExcelDdeFix.exe --repair --yes                    :: 实际修复（需管理员，自动备份）
ExcelDdeFix.exe --rollback backup\dde-backup-xxx.txt --yes
ExcelDdeFix.exe --verify 文件1.xlsx 文件2.xlsx     :: 实测合并验证
ExcelDdeFix.exe --export --html                   :: 导出可视化 HTML 报告
ExcelDdeFix.exe --logs --n 50                     :: 查看最近日志
ExcelDdeFix.exe --version / --help
```

`--diagnose --json` 输出示例（节选）：

```json
{
  "version": "3.0.0",
  "score": 100,
  "verdict": "健康：DDE 合并机制完整",
  "excelExe": "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
  "ddeCleaned": "1",
  "ddeCleanedKind": "DWord",
  "checks": [
    { "progId": "Excel.Sheet.12", "item": "ddeexec", "ok": true,
      "expected": "[open(\"%1\")]", "actual": "[open(\"%1\")]" }
  ]
}
```

> 程序编译为 Windows GUI 子系统。从 `cmd` 运行时会挂回父控制台显示输出；每次调用也都会在 `logs\cli-时间戳.txt` 留一份副本。

---

## 5. 评分规则

健康评分满分 100，逐项校验**值内容**：

| 项目 | 权重 | 判定 |
|---|---|---|
| `ddeexec` 默认值 | 6.25 × 4 | 必须存在且等于 `[open("%1")]` |
| `ddeexec\application` | 6.25 × 4 | 必须以 `EXCEL.EXE` 结尾 |
| `ddeexec\topic` | 6.25 × 4 | 必须等于 `System` |
| `shell\open\command` | 6.25 × 4 | 必须含 `/dde` 且指向探测到的 Excel |

`DDECleaned` / `DDEAllowed` / `DontUpdateLinks` / `IgnoreOtherApplications` / 组策略项**独立监测、不计入评分**。

---

## 6. 命令行参考

| 命令 | 作用 | 需要管理员 |
|---|---|---|
| `--diagnose` | 文本诊断 | 否 |
| `--diagnose --json` | JSON 诊断（可脚本解析） | 否 |
| `--backup [--label 文字]` | 备份当前注册表状态 | 否 |
| `--repair --dry-run` | 演练，列出将写入的内容 | 否 |
| `--repair --yes` | 实际修复（自动备份） | 是 |
| `--rollback 备份文件 --yes` | 回滚到指定备份 | 是 |
| `--verify 文件1 文件2` | 实测打开两个文件验证合并 | 否 |
| `--export --html` | 导出 HTML 报告 | 否 |
| `--logs [--n 行数]` | 查看最近日志 | 否 |
| `--version` / `--help` | 版本 / 帮助 | 否 |

---

## 7. 日志与排障

日志位于运行目录下的 `logs\`：

| 文件 | 内容 |
|---|---|
| `excel-dde-YYYYMMDD.log` | 人类可读：`时间(毫秒) \| 级别 \| [操作] \| 明细` |
| `excel-dde-YYYYMMDD.jsonl` | 机器可读：每行一个 JSON，含 `ts` / `sid` / `level` / `action` / `detail` |
| `cli-YYYYMMDD-HHMMSS.txt` | 每次命令行调用的输出副本 |

**日志级别**：

- `CLICK` / `DONE`：按钮点击的开始与结束（含前后状态、耗时毫秒）
- `CHOICE`：每个确认框你点了 OK / Cancel / Yes / No
- `BACKUP` / `REPAIR` / `ROLLBACK` / `HARDEN` / `REGSVR`：具体操作明细
- `DIFF`：**注册表逐项变化**（`~ key: 旧值 [类型] -> 新值 [类型]`）
- `VERIFY`：实测验证每一步进度与结论
- `AUTO`：自动刷新只在状态变化时才记，不刷屏
- `ERROR`：异常类型、消息、Source、TargetSite、完整堆栈、内部异常

**排障技巧**：

```bat
findstr "\[按钮名\]" logs\excel-dde-*.log     :: 看某次点击的完整链路
:: 有 CLICK 没 DONE = 操作中途异常，往下找同前缀 ERROR
:: 按 sid 过滤 JSONL = 只看某一次程序运行的记录
jq 'select(.level=="DIFF")' logs\*.jsonl       :: 抽出所有注册表变更
```

---

## 8. 编译

**环境**：Windows + .NET Framework 4.0 的 `csc.exe`（随系统自带，无需安装 Visual Studio 或 NuGet）。

```bat
build.bat
```

源码保持 **C# 5 语法兼容**，以便用系统自带的旧编译器直接构建。详见 [CONTRIBUTING.md](CONTRIBUTING.md)。

**目录结构**：

```
xlsx-Excel-GitHub/
├── ExcelDdeFix.exe          # 主程序（也可从 Release 下载）
├── build.bat                # 编译脚本
├── src/
│   ├── RegEngine.cs         # 注册表读写、探测、备份、修复、回滚、评分、报告数据
│   ├── ExcelDdeGui.cs       # 界面、交互、日志引擎
│   ├── Cli.cs               # 命令行入口、实测合并验证、退出码
│   └── HtmlReport.cs        # 自包含 HTML 报告生成
├── examples/                # 脱敏后的示例产物（文档展示用）
│   ├── 示例诊断报告.html
│   ├── 示例操作日志.txt
│   └── 示例日志片段.txt
├── 使用说明.txt             # 中文快速上手卡
├── LICENSE / README.md / CHANGELOG.md / CONTRIBUTING.md / SECURITY.md / .gitignore
└── logs/ backup/            # 运行时生成（已加入 .gitignore，不入库）
```

---

## 9. 安全说明

- 本工具会修改 `HKEY_CLASSES_ROOT` 下的文件关联与 Excel 选项，**修复 / 回滚需要管理员权限**。
- 每次写操作前**自动备份**；备份带 SHA256，回滚前校验完整性。
- `/regserver` 会关闭所有 Excel 进程：先尝试正常关闭窗口，仅对未响应者强制结束。**请先保存工作**。
- 本工具**不收集、不上传任何数据**，无遥测、无网络请求。
- 详见 [SECURITY.md](SECURITY.md)。

---

## 10. 已知限制

- 面向 Windows + Microsoft Excel 2010 及以上；WPS / LibreOffice 的关联会被检测到，但不在修复范围内。
- "Excel 进程 / 工作簿"摘要在评分卡里按"有窗口的进程数"统计；精确验证请用「实测合并验证」，它用 Win32 枚举真实窗口。
- 评分针对 4 个 Excel ProgId；`Word.Document.12` / `PowerPoint.Show.12` 仅体检不修改。

---

## 11. 参与贡献

欢迎提交 Issue 和 Pull Request。提交前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)，重点：

- 保持 C# 5 语法兼容（用系统自带旧编译器构建）
- 注册表写入"先备份后写"，新目标同步更新备份 / 回滚范围
- **报 Issue 或提 PR 时请脱敏**：贴日志前把机器名、用户名、私有路径替换掉（可参考 `examples/` 的脱敏方式）
- 不要在 PR 中包含 `logs/`、`backup/`、`diagnose-*` 等运行产物

如果这个项目帮你解决了问题，欢迎点个 ⭐ Star，让更多人看到。

---

## 12. 许可证

MIT License，详见 [LICENSE](LICENSE)。
