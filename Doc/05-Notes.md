# iWorkHelper 开发注意事项与常见问题

> 文档版本：v0.1.0  ·  更新日期：2025-11-06  ·  适用范围：Outlook VSTO 加载项（VB / .NET Framework 4.8）

## 0. 相关文档
- 开发框架与技术架构设计：[`Doc/01-Architecture.md`](01-Architecture.md)
- 技术栈与兼容性：[`Doc/02-TechStack.md`](02-TechStack.md)
- 开发流程与规范：[`Doc/03-DevelopmentProcess.md`](03-DevelopmentProcess.md)
- 开发任务 TodoList：[`Doc/04-TodoList.md`](04-TodoList.md)
- 变更记录：[`Doc/CHANGELOG.md`](CHANGELOG.md)

## 1. 已知技术限制
- 文本提取范围
  - PdfPig 对文本型 PDF 提取效果稳定；对扫描件或图像型 PDF 能力有限，需引入 OCR 扩展（可选）。
- 附件类型
  - 仅处理本地实际 PDF 附件；链接型/云附件需先下载。
- PDF 合并边界
  - iTextSharp 合并基于页序；对特定加密/增值税电子发票特殊版式可能需额外处理。
- 文件系统权限
  - 归档目录需具备写入权限；网络盘/同步盘（如 OneDrive）可能带来并发与锁定问题。

## 2. 特殊业务逻辑要求（滴滴文档）
- 识别规则（示例关键词）
  - 发票：`"滴滴出行电子发票"`、`"发票号码"`、`"开票日期"`、`"价税合计"`
  - 行程单：`"行程单"`、`"订单号"`、`"车辆使用时间"`、`"起点/终点"`
- 合并策略
  - 同一邮件内优先按日期与金额进行配对；若无法匹配，则仅保留各自独立文件。
  - 合并顺序：发票在前、行程单在后；输出一个合并后的 PDF。
- 重命名规范（建议）
  - `YYYYMMDD_滴滴_姓名_金额_类型.pdf`（类型为发票/行程单/合并）。

## 3. 性能优化建议
- IO 与并发
  - 附件保存与 PDF 读写分离；批量处理时限制并发，避免磁盘争用。
- 文本提取
  - 优先提取关键页（如首页）；对超大 PDF 可采用页级流式读取。
- 合并与重命名
  - 合并前先校验页数与关键字段；重命名基于缓存的抽取结果，避免重复抽取。
- 日志与监控
  - 仅在关键路径写日志；为每次处理生成上下文 ID，便于故障定位。

## 4. 常见问题（FAQ）与解决方案
- 加载项未加载 / 消失
  - 检查 Outlook → 选项 → 加载项，确认状态为「已加载」。
  - 复核注册表：`HKCU\Software\Microsoft\Office\Outlook\Addins\iWorkHelper` 的 `LoadBehavior=3`。
- 发布安装失败（ClickOnce）
  - 证书未信任或过期：更新并重新签名；在信任中心添加为受信任发布者。
- 文本提取为空
  - 文件为扫描版：启用 OCR 扩展或提示用户上传文本型 PDF。
- 合并后顺序错误
  - 修正识别规则；强制发票在前、行程单在后。
- 归档失败（路径不可写）
  - 校验 `ArchivePath` 权限；避免写入同步中的目录；增加重试与后缀策略。

## 5. 命令与终端建议
- 遵循项目规则：涉及命令行操作优先使用 `CMD`。
- 示例（以 `nuget.exe` 配置源为例）：
```cmd
nuget.exe sources Add -Name CorpFeed -Source https://nuget.corp.local/v3/index.json
nuget.exe restore iWorkHelper.sln
```

## 6. 安全与合规
- 隐私：不将邮件内容与附件上传至外部服务；仅在本地处理。
- 许可证：确保 iTextSharp 版本与许可证合规；企业环境需法务评估。

---

> 术语说明请参阅：《开发框架与技术架构设计》中的「附录 A：术语表」（[`Doc/01-Architecture.md`](01-Architecture.md)）。