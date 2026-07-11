# 关键设计决策与风险 DECISIONS_AND_RISKS

> 更新日期：2026-07-11
> 记录采用当前方案的原因、被否决的方案，以及仍然有效的风险。原始接手期风险登记见 [../archive/RISK_REGISTER.md](../archive/RISK_REGISTER.md)（历史快照，多数已解决）。

## 一、关键设计决策

| 决策 | 选择 | 原因 / 被否决方案 |
|------|------|-------------------|
| 目标框架 | .NET Framework 4.8（VSTO 传统模型） | 工程实际配置；`.trae` 曾写".NET Framework 8.0"不成立，以工程为准 |
| PDF 文本抽取 | PdfPig 0.1.14 | 文本型 PDF 足够；图片型走百度 OCR。PdfPig **非 OCR** |
| PDF 合并 | PdfPig 自带 `PdfMerger` | **不引入新库**、无新增许可证风险 |
| HTTP/JSON | 内置 `HttpWebRequest` + `JavaScriptSerializer` | 不引第三方库，降低依赖与体积 |
| OCR 时机 | 合并前分别识别、再 `InvoiceRecognitionMerger` 合并结果 | 各 PDF 在原始上下文识别质量更好，无需依赖多页 OCR |
| 本地行程单解析 | PdfPig 词坐标按列取单元格 | 否决"全文拼接+正则切分"（对终点无 `区\|`、长地址换行、车型别名不稳）|
| 常规发票识别 | 专用 `GeneralInvoiceLocalRecognizer`（委派） | 否决与滴滴共用一套正则（购销混淆、误取明细金额、缺分区/评分）|
| 命名规则 | 单一统一模板 + fallback | 否决发票/行程/未识别三套模板（简化配置，旧三套仅兼容）|
| 密钥保护 | DPAPI（CurrentUser） | .NET 内置、无第三方依赖；不追求强多方安全（见下风险）|
| 运行锁 | `ArchiveRunGuard`（`Interlocked` 原子 + `Using`） | 否决"预检查内维护运行标志"（导致自我阻断缺陷）|
| 进度刷新 | UI 线程 + `Application.DoEvents` | 后台线程访问 Outlook COM 有跨线程风险；暂不引入 BackgroundWorker |
| 内外网版本 | 编译期开关 `BuildFeatures` 门控在线 OCR | 内网合规要求禁止外网请求；默认（未定义 `INTERNET_BUILD`）即禁用 |

## 二、不可违背的约束

- **内网版禁止任何在线 OCR / 外网请求**：统一通过 `BuildFeatures.OnlineParserEnabled` 门控，见 [../architecture/BUILD_VARIANTS.md](../architecture/BUILD_VARIANTS.md)。
- **不硬编码密钥**：AK/SK 仅来自设置/加密配置/环境变量；日志、文档、弹窗**不得出现明文 AK/SK/token**（仅脱敏摘要）。
- **样例 PDF、真实 OCR 返回、`local-debug/` 诊断、`user.config`、`baidu-ocr.config.xml`、私钥不入库**（`.gitignore` 覆盖）。
- **尽量不丢文件**：识别失败也以回退名归档；仅"复制动作本身"失败才记 Failure。

## 三、仍然有效的风险

| 风险 | 说明 | 应对 |
|------|------|------|
| PdfPig net471 DLL 在 net48 工程 | 构建与本地运行已验证；极端环境仍需留意 | 如遇加载异常，改用 net48/netstandard2.0 DLL 或升级 PdfPig |
| `Option Strict Off` | 类型宽松，业务代码易埋运行期类型错误 | 新增关键转换显式处理；可逐文件评估开启 `Option Strict On` |
| Outlook COM 陷阱 | 多选/非邮件项/COM 未释放导致泄漏或卡死 | `MailAttachmentReader` 已严格判类型 + `Marshal.ReleaseComObject`；改动时保持 |
| UI 线程同步归档 | 大批量/大 PDF 可能短时假死 | 当前 `DoEvents` 刷新；若假死明显再评估后台化 |
| DPAPI 安全边界 | 仅保护"当前用户+当前机器"，换用户/机器需重输 SK | 属预期；不替代访问控制 |
| 临时签名证书 | 发布用临时 pfx | 发布前替换正式代码签名证书 |
| Outlook 启动被 Resiliency 禁用 | 已优化 Startup，待真实环境验证 | 见 [../deployment/STARTUP_PERFORMANCE.md](../deployment/STARTUP_PERFORMANCE.md) |
| Office/Outlook 版本差异 | Interop 15.0 与目标 Office 365 差异 | `EmbedInteropTypes` 已缓解；目标版本实测 |

## 四、已解决（原接手期风险）

无日志/无异常处理、非 Git 仓库、核心业务缺失、OCR 方案未定、命名规范未定、缺样例等，均已在后续开发中解决（见 [DEVELOPMENT_HISTORY.md](DEVELOPMENT_HISTORY.md)）。
