# 更新日志

本项目遵循语义化版本（SemVer）思路。

## v3.0.0 — 2026-11-22

开源就绪版。从"能修"升级为"可审计、可自动化、可验证"。

### 新增
- **精确评分**：从"检查键是否存在"升级为逐项校验值内容（4 ProgId × 4 项 × 6.25 分）
- **智能 Excel 探测**：App Paths → COM LocalServer32 → 多版本多路径，带存在性校验；找不到即中止，避免写错路径
- **多版本 Office**：覆盖 16.0 / 15.0 / 14.0 的 Excel 选项，备份与报告全覆盖
- **实测合并验证**：GUI 按钮 + CLI `--verify`，自动打开两个 xlsx、统计新增进程、只关闭自己拉起的进程
- **命令行接口**：`--diagnose` / `--backup` / `--repair` / `--rollback` / `--verify` / `--export` / `--logs` / `--version` / `--help`，含 JSON 输出与退出码
- **演练模式**：`--repair --dry-run` 无需管理员权限即可预览将写入的内容
- **JSONL 结构化日志**：会话 ID、逐项 diff、验证进度落盘
- **HTML 可视化报告**：评分趋势 SVG、事件时间线、环境指纹、逐项校验
- **备份 SHA256 校验 + 回滚保护快照 + 危险删除护栏**
- **单实例 Mutex**，防止并发写注册表竞态
- **环境指纹**：OS / .NET / Excel 版本 / PE 位数 / Office 版本 / 权限
- 开源配套：LICENSE / README / CONTRIBUTING / SECURITY / .gitignore

### 修复
- 修正 `--verify` 把第一个文件误当第二个参数的解析错误
- 修正 `--dry-run` 被权限检查误挡的问题
- `/regserver` 改为先 `CloseMainWindow()` 再按需 `Kill()`，并加强未保存工作警告
- 评分不再因安全加固标记 `DDECleaned=1` 虚降

---

## v2.0.0 — 2026-11-22

修正安全项认知后的版本。

- 新增「DDE 安全加固监测」面板，独立展示 8 项安全设置
- 评分规则改为 ddeexec 50 + command 50，`DDECleaned` 不再扣分
- 备份 / 读写 / 回滚按注册表原类型（String / DWord / Missing）处理
- 新增「修正安全项类型」按钮，把被写坏的 `DDECleaned` 纠正为 `DWord 1`
- 日志升级为结构化：CLICK / DONE / CHOICE / 异常堆栈 / 前后状态快照 / 耗时
- 诊断报告纳入安全监测段

---

## v1.0.0 — 2026-11-22

初版。

- 诊断评分（当时把 `DDECleaned` 计 20 分）
- 状态表格、自动刷新
- 备份 / 一键修复 / 回滚 / `/regserver` / 管理员重启 / 诊断报告
- 基础日志、托盘式浅色 WinForms 界面
