# 开发进度 CURRENT_PROGRESS

> 历史说明：本文生成于 2026-07-06，反映的是接手初期状态。当前仓库已继续演进，本文中的完成度与未实现项不再代表最新事实；请结合根目录 `README.md`、`docs/README.md` 与后续 `PHASE_*` 文档阅读。

> 生成日期：2026-07-06
> 说明：所有条目均给出代码/配置依据。无依据者标注“待确认”。

## 汇总

| 分类 | 数量 | 代表项 |
|------|------|--------|
| ✅ 已完成 | 3 | VSTO 工程骨架、Ribbon UI、设置界面 |
| 🟡 部分完成 | 2 | 归档按钮（仅占位）、解析模式设置（仅开关无消费方） |
| ❌ 未开始 | 6 | 附件读取、PDF 解析、OCR、字段提取、重命名、归档 |
| ❓ 待确认 | 4 | OCR 具体方案、在线服务、滴滴解析规则、目标框架说明 |

**整体完成度估计：约 10%~15%（仅 UI 与设置层）。**

---

## 一、✅ 已完成

### 1. VSTO Outlook 插件工程骨架
- 依据：`iWorkhelper.vbproj`（`ProjectTypeGuids` 含 VSTO GUID `BAA0C2D2-...`；`OutputType=Library`；`TargetFrameworkVersion=v4.8`；`OfficeApplication=Outlook`）。
- `ThisAddIn.vb`：定义了 `ThisAddIn` 类与 `Startup`/`Shutdown` 事件处理（**方法体为空**，`ThisAddIn.vb:3-9`）。
- 结论：工程可作为 Outlook 加载项加载，但启动/关闭时无任何逻辑。

### 2. 功能区（Ribbon）UI
- 依据：`MainRibbon.Designer.vb`。
- 已定义：Tab `工作助手`（`TabWorkHelper`）、Group `发票归档`（`GroupInvoiceFiling`）、按钮 `归档`（`ButtonArchive`）与 `设置`（`ButtonSettings`）。
- RibbonType：`Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read`（`MainRibbon.Designer.vb:83`）——即在阅读邮件/资源管理器场景显示。

### 3. 设置界面（SettingsForm）
- 依据：`SettingsForm.vb`（完整可用逻辑）。
- 功能：
  - 归档目录选择（`btnBrowseFolder_Click`，`SettingsForm.vb:23-30`，用 `folderBrowserDialog1`）。
  - 归档目录持久化（`SaveFolderPath` → `My.Settings.ArchiveFolderPath`，`SettingsForm.vb:54-61`）。
  - 解析模式单选切换 Local/Online（`rdoLocalParse`/`rdoOnlineParse`，`SettingsForm.vb:36-52`）→ `My.Settings.ParseMode`。
  - 版本号显示（读取程序集版本，`SettingsForm.vb:16-17`）。
- 设置由 `ButtonSettings_Click` 以模态对话框打开（`MainRibbon.vb:19-23`）。
- 结论：这是当前**唯一功能完整**的模块。

---

## 二、🟡 部分完成

### 4. 归档按钮（ButtonArchive）——仅占位
- 依据：`MainRibbon.vb:9-17`。
- 现状：点击后仅读取 `My.Settings.ParseMode`，弹出 `MsgBox("在线模式")` 或 `MsgBox("本地模式")`。
- 缺失：无附件读取、无解析、无归档。属于**流程占位/自测代码**。

### 5. 解析模式设置（ParseMode）——有开关无消费方
- 依据：设置项定义 `My Project\Settings.Designer.vb:66-74`（默认 `Local`）；写入方 `SettingsForm.vb`；读取方仅 `MainRibbon.vb:10`（仅用于弹框）。
- 现状：Local/Online 可切换并持久化，但**没有任何真实解析逻辑消费该模式**。

---

## 三、❌ 未开始（核心业务，全部缺失）

> 以下均为“全仓库搜索无对应实现”，依据为 `Core\Local OCR\`、`Core\Online OCR\` 均为**空文件夹**（`iWorkhelper.vbproj:303-306` 仅登记了空 Folder），且 `.vb` 源码中无相关方法。

| # | 功能 | 缺失说明 | 依据 |
|---|------|---------|------|
| 6 | 邮件附件读取 | 无任何 `Outlook.MailItem` / `Attachments` 访问代码 | 全仓库无相关 API 调用；`ThisAddIn.vb` 未持有 Application 引用逻辑 |
| 7 | PDF 解析 | 引用了 PdfPig 但**代码中从未使用** | `iWorkhelper.vbproj:141-161` 引用 PdfPig；无 `Imports UglyToad`、无 `PdfDocument.Open` 调用 |
| 8 | OCR 文本识别 | 无 OCR 引擎/服务；仅两个空文件夹占位 | `Core\Local OCR\`、`Core\Online OCR\` 为空 |
| 9 | 字段提取（发票号/金额/日期/行程/滴滴等） | 无任何解析规则或正则 | 全仓库无相关代码或样例 |
| 10 | 文件重命名 | 无 `File.Move`/命名规则 | 全仓库无相关代码 |
| 11 | 归档/分类落盘 | 设置了 `ArchiveFolderPath` 但无写入逻辑 | 仅设置项存在，无消费方 |

---

## 四、❓ 待确认

1. **OCR 具体方案**：是本地 OCR 引擎（如 Tesseract/PaddleOCR/Windows.Media.Ocr）还是仅用 PdfPig 抽取文本？“Local OCR”命名与 PdfPig（非 OCR）能力不一致。**待确认**。
2. **在线解析服务**：`Online` 模式对接哪个云 OCR/发票识别服务（如百度/腾讯/阿里发票识别、滴滴开放平台等）？是否有 API Key/账号？**待确认**。
3. **滴滴发票/行程单解析规则**：字段清单、命名规范、分类规则、样例 PDF 是否可提供？**待确认**。
4. **目标框架说明差异**：`.trae` 规则写“.NET Framework 8.0”，实际工程为 .NET Framework 4.8。以哪个为准？（本文档按工程实际 4.8 判断）**待确认**。

---

## 五、其他事实

- **非 Git 仓库**：目录下无 `.git`（`git status` 报 `not a git repository`）。因此无提交记录、无分支、无未提交变更可分析。**建议尽快 `git init` 纳入版本管理**（见风险登记）。
- **无 README / 设计文档 / 需求文档 / 变更记录**：仓库根目录及子目录未发现任何既有说明文档（`.trae/rules/projectIntroduction.md` 仅 4 行环境说明）。
- **无 TODO/FIXME 标记**：源码中未发现 `TODO`/`FIXME` 注释；未完成部分体现为空方法体与空文件夹，而非注释标记。
- **存在临时签名证书**：`iWorkhelper_TemporaryKey.pfx`（VSTO 默认临时清单签名，`iWorkhelper.vbproj:324` `ManifestKeyFile`）。发布前需替换为正式证书（见风险登记）。
