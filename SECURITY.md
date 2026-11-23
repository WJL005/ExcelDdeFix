# 安全说明

## 支持范围

- 当前支持版本：v3.x
- v1 / v2 已停止维护（v1 曾用错误类型写 `DDECleaned`，v2 未做备份哈希校验）

## 本工具会做什么

- 读取并在修复时写入 `HKEY_CLASSES_ROOT` 下 4 个 Excel ProgId 的 `shell\open\command` 与 `ddeexec` 相关值
- 读取 Excel Options 与安全策略项；仅在必要时把被写坏的 `DDECleaned` 类型纠正为 `DWord 1`（保持"已加固"语义）
- `/regserver` 会结束所有 Excel 进程并重新注册 COM/DDE

**不会**联网、不会上传数据、不会修改用户文档、不会安装服务或驱动。

## 使用安全建议

1. 修复前先点「备份当前状态」；工具在修复前也会自动备份
2. 确认备份文件带 `v3(sha256=...)` 标记；回滚前工具会校验哈希，校验失败会拒绝回滚
3. 执行 `/regserver` 前保存并关闭所有 Excel 中未保存的工作
4. 回滚操作会在执行前再留一份「回滚前保护快照」，可二次回退
5. 从公共位置下载 exe 时请核对哈希；本仓库不提供代码签名，首次运行可能触发 SmartScreen 提示

## 备份文件格式

```
# Excel DDE 修复工具 - 状态备份 v3
at=2026-11-22 16:20:00
label=修复前自动备份
score=100
# 格式: PATH<TAB>VALUENAME<TAB>KIND<TAB>VALUE
HKEY_CLASSES_ROOT\Excel.Sheet.12\shell\open\command			String	"C:\...\EXCEL.EXE" /dde "%1"
...
# sha256=xxxxxxxx...
```

`KIND=Missing` 表示该值原本不存在，回滚时会被删除。出于安全，工具**只允许删除以 `\ddeexec` 结尾的键树**，拒绝删除其他整键，防止被篡改的备份造成破坏。

## 报告安全漏洞

请通过仓库 Issue 私聊维护者，或邮件联系。报告请包含：

- 问题类型（越权写入 / 备份校验绕过 / 命令注入等）
- 复现步骤与最小 PoC
- 影响版本

请勿在公开 Issue 里直接贴可造成破坏的备份文件内容。

## 隐私

- 诊断报告包含机器名、用户名、Excel 路径。**分享报告前请自行检查并脱敏**。
- 日志同样包含这些环境信息，仅保存在本地 `logs\` 目录，工具不会主动发送。
