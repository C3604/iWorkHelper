# 常规发票本地识别 GENERAL_INVOICE_RECOGNITION

> 更新日期：2026-07-11
> 由「现状分析 + 优化报告 + 真实样例验证」整合。回归方法见 [../testing/REGRESSION_GUIDE.md](../testing/REGRESSION_GUIDE.md)。
> 范围：仅本地识别（非 OCR）。不改动滴滴发票、滴滴+行程单合并、在线百度 OCR 链路。

## 1. 背景与根因（为什么需要专用识别器）

改造前，常规发票（非滴滴增值税发票）由 `LocalTextInvoiceRecognizer` 处理，**与滴滴发票共用同一套滴滴校准正则**，导致识别率低。根因：

1. **滴滴校准正则不通用**：`AssignParties` 的名称正则要求名称后紧跟"统一社会信用代码/纳税人识别号"，数电票自然人购买方常无税号，导致名称抽取失败或购销错配。
2. **二维版式被一维 `page.Text` 破坏**：购销是左右并排两栏、商品明细是多列表格，`page.Text` 拼成一维时顺序不稳；已有坐标行重建当时只服务滴滴行程明细。
3. **首个匹配即采用，无候选评分**：金额在发票中多次出现（明细金额/合计/税额/价税合计），`MatchFirst` 可能误取明细金额；发票号码也可能撞订单号/税号。
4. **命名充分度不含销售方**：常规发票命名模板需销售方名称，但 `IsNamingSufficient` 只看"日期+金额"，缺销售方仍判充分 → 该回退 OCR 时没回退。
5. **无商品明细结构**：`InvoiceInfo` 无 `LineItems`。

## 2. 实现方案

`LocalTextInvoiceRecognizer` 检测到**非滴滴的 VatInvoice** 时委派 `GeneralInvoiceLocalRecognizer`（滴滴/行程单逻辑完全不变）：

```
GeneralInvoiceLocalRecognizer.Recognize
  1) 归一化文本 + 构建「逻辑行」（有坐标用坐标行重建，无坐标按换行切分）
  2) 识别发票类型（专票/普票/数电票/电子发票…；仅诊断与命名参考，不改分流）
  3) 逐字段生成候选并评分择优（GeneralInvoiceCandidateScorer）
  4) 商品明细解析（PdfTableRegionDetector + GeneralInvoiceLineItemParser，失败不影响头部）
  5) 状态：开票日期+金额+销售方齐全→Success；缺→PartialSuccess（回退 OCR）；无核心→NeedsOcr
```

**新增/改动模块**：`GeneralInvoiceLocalRecognizer`、`GeneralInvoiceFieldCandidate`、`GeneralInvoiceParseResult`、`GeneralInvoiceCandidateScorer`、`GeneralInvoiceLineItemParser`、`InvoiceLineItem`、`PdfTextBlock`、`PdfTableRegionDetector`；`KeyFieldEvaluator` 新增 `GetMissingGeneralInvoiceNamingFields`（开票日期+金额+销售方）。

## 3. 字段候选与评分要点

同一字段收集多个候选，按「标签就近 + 区域合理 + 格式合理 − 干扰惩罚」打分取最高分：

| 字段 | 策略 | 防误取 |
|------|------|--------|
| 开票日期 | 靠近「开票日期」标签、头部区、格式完整（`yyyy-MM-dd`/`/`/`.`/`年月日`） | 不取明细/行程日期 |
| 发票号码 | 靠近「发票号码」标签；长度 20（数电）/8（传统）最高，10–12 降分；号码后不得紧跟字母 `(?![0-9A-Za-z])` | 不从「发票代码」行取；行内含税号/校验码/订单号时降分 |
| 发票代码 | 靠近「发票代码」标签，10/12 位 | 数电票允许为空 |
| 销售方名称 | 靠近「销售方/销方」标签或区域 + 组织特征词（有限公司/公司/经营部/店…） | 不被购买方/项目名称覆盖 |
| 购买方名称 | 靠近「购买方/购方」；与销售方按 X 坐标（有坐标）或行序（无坐标）分栏 | 购销不混淆 |
| 价税合计（{金额}） | 价税合计(小写) > 全局价税合计后首金额；同行取「小写」侧最后一个两位小数 | 税额/税率降分；**不取明细行金额** |
| 税率 / 税额 | `13/9/6/3/1/0%`、免税、不征税 / 「合计…税额」区金额 | — |

命名 `{金额}` 对常规发票取 **价税合计 > 金额**（无行程金额），与择优一致。字段清单见 [../requirements/FIELD_EXTRACTION_SPEC.md](../requirements/FIELD_EXTRACTION_SPEC.md)。

## 4. 商品明细解析

`GeneralInvoiceLineItemParser` 在 `PdfTableRegionDetector` 检出的明细区（表头「项目名称/货物+金额/税额」→「合计/价税合计」之间）逐行尽力解析：名称优先 `*类别*名称`、非 `*` 且几乎无数字的行视为名称换行续行；行内多个两位小数时末位=税额、其前=金额。单行失败跳过，明细失败不影响头部字段；明细金额+税额与价税合计不一致仅记提示。结果写入 `InvoiceInfo.LineItems`。

## 5. OCR 回退（关键修复）

常规发票缺 开票日期/金额/销售方名称 任一 → 识别器返回 `PartialSuccess`，`RecognitionPipeline` 据 `IsNamingSufficient=False` 在**启用且配置完整**时回退百度 OCR（OCR 成功优先其结构化结果，本地也有字段则 Mixed）；OCR 未启用/未配置 → 部分成功 + 用户可读原因，命名走 fallback。

## 6. 验证结果

- **合成自测（真实执行）**：OfflineTester `--selftest [8]` 覆盖 委派、购销不混淆、`{金额}` 取价税合计而非明细金额、发票号码不误取税号/校验码、开票日期多格式、商品明细（含多条/名称换行）、缺销售方→回退 OCR、滴滴不受影响。全量 selftest 通过。
- **真实样例（3 份脱敏常规发票 PDF）**：修复销售方名称识别（`CleanName` 清行尾"购/销"字、`ParseParties` 优先匹配"销 名称:/购 名称:"明确模式）后，开票日期/金额/销售方/发票号码/明细行数 3/3 命中，命名核心字段完整度 3/3。
  - 命名预览（脱敏示例）：`20260623_24.00_示例物流有限公司.pdf`、`20260623_950.00_示例航空股份有限公司.pdf`。
  - 运单类无发票字段 → 判定字段不足、建议回退 OCR（符合预期）。
  > 真实样例 PDF 与期望值均为本地忽略、不入库（`.gitignore` 覆盖 `sample/*.pdf`）。

## 7. 建议在线 OCR 的场景 / 遗留

- 建议 OCR：图片型/扫描件 PDF、版式特殊/购销区非标注、明细跨页复杂表格、手写/异常字体。
- 遗留：真实票面大规模校准（当前阈值基于合成用例）；明细列坐标级精确分列（规格/单位/数量/单价常留空）；大写金额→数字换算（当前仅保留原文）。
