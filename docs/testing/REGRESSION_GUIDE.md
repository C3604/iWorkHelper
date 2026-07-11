# 本地识别回归指南 REGRESSION_GUIDE

> 更新日期：2026-07-11
> 由「本地解析回归」+「常规发票回归」整合。工具用法见 [OFFLINE_TESTER_GUIDE.md](OFFLINE_TESTER_GUIDE.md)。
> **安全约束**：真实/脱敏 PDF 样例、期望文件、`local-debug/` 诊断输出**一律不提交**（`.gitignore` 已覆盖 `sample/*.pdf`、`local-debug/`、`*.localdebug.txt`）。

---

## 一、滴滴行程单本地回归（命名核心字段）

回归验证统一命名模板的 4 个核心字段：`{乘车日期}`（行程出发日期>开票日期）、`{金额}`（行程金额>价税合计>金额）、`{出发地点}`（起点）、`{到达地点}`（终点）。

```
:: 5 份样例本地识别 + 命名预览
tools\OfflineTester\bin\Debug\OfflineTester.exe sample --local-only

:: 逐份诊断（原始/归一化文本、坐标行、字段、命名、缺失、是否需 OCR → local-debug\，忽略目录）
tools\OfflineTester\bin\Debug\OfflineTester.exe sample --local-only --dump-local-debug
```

**判定**：单份"通过"= 4 个核心字段均非空且与期望一致（地点可用关键词包含匹配，避免 `|`→`_` 清理差异误判）；某字段为空 → 记为字段不足、应由在线 OCR 兜底。

**已知结果**：现有 5 份真实样例本地 4 核心字段 5/5 命中（见 [../development/LOCAL_TEXT_RECOGNITION_DIAGNOSTIC_REPORT.md](../development/LOCAL_TEXT_RECOGNITION_DIAGNOSTIC_REPORT.md)）；图片型/异版式走在线 OCR。

---

## 二、常规发票本地回归

设计见 [../development/GENERAL_INVOICE_RECOGNITION.md](../development/GENERAL_INVOICE_RECOGNITION.md)。

### 2.1 快速诊断单张发票

```
OfflineTester.exe sample\general-invoice\a.pdf --general-invoice --dump-layout --dump-candidates --dump-line-items
```

输出：PDF 分类、发票类型、逐页行（页/行/坐标/文本）、字段候选与评分、**候选评分前 3 名（命名核心字段）**、择优结果、**本地字段完整度 N/3**、商品明细、缺失、**是否建议回退 OCR**、常规发票命名预览。候选前 3 名与完整度始终输出（无需 dump 参数）。

**校准要点**：先看「候选评分前 3 名」判断是否择优错误（金额取到明细金额、号码取到税号）；若某字段「(无候选)」说明抽取规则未命中该版式，需在 `GeneralInvoiceLocalRecognizer` 补候选规则；择优错误则在 `GeneralInvoiceCandidateScorer` 调权重。

### 2.2 期望文件回归（CSV/JSON，按 PDF 文件名匹配）

期望文件放 `sample/general-invoice/expected/`（忽略目录）。至少含：PDF 文件名、期望开票日期、期望金额、期望销售方；可选期望发票号码、明细行数。

CSV 示例（列名兼容中英文 `pdf/文件名`、`date/开票日期`、`amount/金额`、`seller/销售方名称`、`number/发票号码`、`items/明细行数`；空列=不校验）：
```
pdf,date,amount,seller,number,items
a.pdf,2026-05-20,106.00,示例卖方有限公司,12345678901234567890,1
```

```
OfflineTester.exe sample\general-invoice --general-invoice --compare-expected sample\general-invoice\expected\expected.csv
```

输出每字段 `PASS/FAIL`（含实际/期望值）+ 汇总「字段通过 N，失败 M」；有失败退出码 1。金额经 `NormalizeAmount`（忽略 ¥/元/千分位）、日期经 `NormalizeDateToYmd`（多格式统一 8 位）、销售方/号码精确比较。

---

## 三、内置合成自测（随构建回归，无需真实样例）

```
OfflineTester.exe --selftest
```

其中 `[8] 常规发票本地识别` 覆盖：委派、购销不混淆、金额取价税合计而非明细、发票号码不误取、开票日期多格式、商品明细（含数电票消歧/多条明细/名称换行）、缺销售方→回退 OCR、滴滴不受影响。**用例为合成非敏感数据，随代码入库用于防回归。**

---

## 四、反馈校准流程

1. 用 `--general-invoice --dump-candidates`（或 `--dump-local-debug`）跑真实票面，观察择优/字段错误。
2. 仅摘录**脱敏**关键片段（去真实税号/名称/号码）反馈，或用合成数据复现同版式问题。
3. 在 `GeneralInvoiceCandidateScorer` 调权重或在 `GeneralInvoiceLocalRecognizer` 补候选规则。
4. 在 `--selftest [8]` 补一条对应合成用例，保证不回归。
