# 贡献指南

欢迎提交 Issue 和 Pull Request。本工具面向真实排障场景，改动请以"不破坏用户注册表"为第一原则。

## 开发环境

- Windows 7 及以上
- .NET Framework 4.0 的 `csc.exe`（系统自带，路径通常为 `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`）
- 不需要 Visual Studio 或 NuGet

## 编译

```bat
build.bat
```

`build.bat` 会在本机查找 64 位 / 32 位 `csc.exe`，编译四个源文件为 `ExcelDdeFix.exe`。

## 代码约定

1. **保持 C# 5 语法兼容**：不要用字符串插值 `$"..."`、`?.`、`nameof`、表达式体成员、自动属性初始化器。原因是要用系统自带的旧编译器直接构建。
2. **字符串里的引号**：中文说明文字里要用引号时请用中文引号「」或『』，不要用 ASCII 双引号，否则会提前闭合 C# 字符串。
3. **注册表写入必须成对**：先备份再写；新目标要同时加入 `CaptureCurrent()` 和 `Restore()` 的处理范围。
4. **评分规则变更要同步三处**：`RegEngine.Diagnose()`、`HealthReport`/`ExportReport()` 中的说明文字、README 的评分表。
5. **所有用户可见操作都要落日志**：按钮点击由 `MakeButton` 自动包 CLICK/DONE；函数内部用 `Log.WriteEx` 记关键参数和结果。
6. **CLI 命令必须有退出码语义**，并在 `--help` 中说明。

## 测试

- 用只读方式验证：`ExcelDdeFix.exe --diagnose --json`，确认 `score` 与 `checks[]` 符合预期
- 演练修复：`ExcelDdeFix.exe --repair --dry-run`（不写注册表）
- 实测合并：`ExcelDdeFix.exe --verify 文件1.xlsx 文件2.xlsx`
- 任何涉及注册表写入的改动，请先在虚拟机或测试机上验证备份 / 回滚往返一致
- 自测程序请用临时文件，提交前删除，不要留在交付目录

## 提交 PR

1. 说明解决了什么问题，附上日志片段或截图（注意脱敏机器名 / 用户名）
2. 如果是评分规则变更，说明为什么该值影响或不影响文件打开合并，最好附实测证据
3. 更新 `CHANGELOG.md` 与相关文档
4. 不要在 PR 中包含 `logs/`、`backup/`、`diagnose-*` 等运行产物

## 报告问题

请附：

- `--diagnose --json` 的完整输出
- `logs\excel-dde-当天.log` 中相关操作的 CLICK → DONE 链路
- Excel 版本与位数（报告里有）
- 复现步骤

详见 [SECURITY.md](SECURITY.md) 了解安全相关的报告方式。
