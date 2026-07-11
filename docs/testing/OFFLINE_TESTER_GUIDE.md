# 离线测试工具指南 OFFLINE_TESTER_GUIDE

> 更新日期：2026-07-11
> 工具：`tools/OfflineTester`（独立控制台，不依赖 Outlook；链接 Core 离线安全源文件）。
> 回归方法见 [REGRESSION_GUIDE.md](REGRESSION_GUIDE.md)。

## 1. 用途

在不启动 Outlook 的情况下，验证：读取 PDF → PdfPig 文本抽取 → 本地字段解析 →（可选）百度 OCR → 字段映射 → 命名预览 → 摘要。用于快速迭代解析规则与联调百度 OCR。

## 2. 构建

```
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" ^
  tools\OfflineTester\OfflineTester.vbproj /t:Build /p:Configuration=Debug /p:Platform=AnyCPU
```
产物：`tools\OfflineTester\bin\Debug\OfflineTester.exe`

## 3. 用法

```
OfflineTester.exe <pdf文件或目录> [选项]
```

| 选项 | 说明 |
|------|------|
| （无选项） | 仅本地解析 |
| `--local-only` | 仅本地文本解析，不调用在线 OCR |
| `--ocr` | 本地优先，字段不足/图片型时回退在线 OCR |
| `--force-ocr` | 强制调用在线 OCR（即使本地充分） |
| `--save-response <目录>` | 保存**脱敏后**的 OCR 原始响应到指定目录 |
| `--general-invoice` | 常规发票本地识别诊断：分类/发票类型/字段候选/择优/明细/缺失/是否需 OCR/命名预览 |
| `--dump-layout` | 配合 `--general-invoice`：打印逐页行（页/行/坐标/文本） |
| `--dump-candidates` | 配合 `--general-invoice`：打印每字段候选与评分、择优结果 |
| `--dump-line-items` | 配合 `--general-invoice`：打印商品明细 |
| `--compare-expected <文件>` | 常规发票回归对比（json/csv 期望，按 PDF 文件名匹配）；输出每字段 PASS/FAIL 与实际/期望，汇总通过/失败数；有失败退出码 1。**期望文件与真实样例均不提交** |

`--general-invoice` 诊断固定输出（无需 dump 参数）：PDF 分类、发票类型、识别状态/来源、**候选评分前 3 名（发票号码/开票日期/销售方/价税合计）**、择优字段、商品明细行数、**本地字段完整度 N/3**、**是否建议回退在线 OCR**、常规发票命名预览。加 `--dump-layout/--dump-candidates/--dump-line-items` 展开逐页行/全部候选/明细。
| `--template <模板>` | 用自定义**发票**命名模板做命名预览（验收用） |
| `--parse-ocr-json <文件>` | 解析一份（脱敏）百度 OCR JSON，跑通解析器+映射器（不联网） |
| `--selftest` | 内置自测：DPAPI 往返 / 命名模板边界 / 合成 OCR JSON / 识别合并（35 项） |
| `--dump-local-debug` | 转储本地解析诊断到 `local-debug\`（原始/归一化文本、坐标行、字段、命名、缺失、是否需 OCR；**含票据信息、勿提交**） |
| `--classify` | 输出每个 PDF 的分流分类（DidiInvoicePdf/DidiTripPdf/GeneralInvoicePdf/UnknownPdf）及对应命名 |
| `--preflight [目录]` / `--validate-config [目录]` | 归档前预检查/配置校验（目录/OCR/模板/临时/日志）；有阻断项 exit 1 |
| `--simulate-error <码名\|list>` | 展示某错误码的用户友好文案（不含敏感信息） |
| `--save-baidu-config` | 用环境变量 AK/SK 写入本机加密配置（SK 经 DPAPI 加密，Enabled=true） |
| `--use-config` | 使用本机加密配置文件（DPAPI 解密 SK）执行 OCR（不依赖环境变量） |
| `--help` / `-h` | 显示帮助（无参数或参数错误时也会显示） |

### 配置落地闭环（第七阶段，真实执行通过）
```
set BAIDU_OCR_AK=xxx
set BAIDU_OCR_SK=yyy
OfflineTester.exe --save-baidu-config          :: 写入 %AppData%\iWorkHelper\baidu-ocr.config.xml（SK DPAPI 加密）
set BAIDU_OCR_AK=                               :: 清空环境变量
set BAIDU_OCR_SK=
OfflineTester.exe sample\a.pdf --use-config --force-ocr   :: 从加密配置解密并真实 OCR
```
该配置文件同时供 Outlook 加载项（经 OcrConfigProvider）读取。**文件已被 .gitignore 忽略，切勿提交。**

### 退出码
| 退出码 | 含义 |
|--------|------|
| 0 | 正常（含 selftest 全通过） |
| 1 | 运行异常 / selftest 有失败项 |
| 2 | 参数错误 / 指定文件不存在 |

### 在线 OCR 密钥（不写入代码/日志）
```
set BAIDU_OCR_AK=你的APIKey
set BAIDU_OCR_SK=你的SecretKey
```
工具仅打印 AK 的**脱敏**值（如 `ab****yz`），从不打印完整 AK/SK/token。

## 4. 示例

```
:: 仅本地解析整个样例目录
OfflineTester.exe sample

:: 仅本地解析单个文件
OfflineTester.exe sample\2026-05-18_138.46_滴滴出行_合并.pdf --local-only

:: 本地优先 + 在线兜底
set BAIDU_OCR_AK=xxx
set BAIDU_OCR_SK=yyy
OfflineTester.exe sample --ocr

:: 强制在线 OCR 并保存脱敏返回
OfflineTester.exe sample --force-ocr --save-response ocr-raw
```

## 5. 输出说明

每份 PDF 输出：
- 抽取：成功/页数/有效字符/是否疑似图片型
- 识别：状态 / **来源（LocalText / BaiduOcr / Mixed / Failed）** / 类型 / 字段数
- 关键字段（发票号码/开票日期/销售方/购买方/价税合计/税率）
- **行程明细数量**（及声明笔数）与逐条行程（车型/出发/城市/起点/终点/里程/金额）
- 缺失关键字段
- **命名规则 + 命名预览**（默认模板）
- 失败原因（若失败）
- 未知占位符（若模板含）

## 6. 已验证结果（真实执行）

- `--selftest`：全部 PASS（DPAPI 往返、命名模板边界、合成 OCR JSON 解析映射、常规发票本地识别、运行锁、分流等；用例随功能增补，以实际运行输出为准）。
- 5 份真实样例（默认/`--local-only`）：全部 `Success / LocalText`，命名与字段正确，多条行程解析生效，退出码 0。
- `--force-ocr` / `--ocr`（无 AK/SK）：优雅回退本地，提示“未配置 AK/SK”，不崩溃。
- `--parse-ocr-json`：解析合成脱敏 JSON，映射 VAT 字段正确。
- 退出码：错误参数=2、缺文件=2、selftest 全过=0。
- **真实在线 OCR：已用真实 AK/SK 联调成功**——5 份样例 `Success/Mixed/VatInvoice/31 字段`；联调中修复 `Uri.EscapeDataString` 超长 bug（改 `WebUtility.UrlEncode`）；日志无 AK/SK/token 明文。详见 [../development/OCR_ONLINE_VALIDATION_LOG.md](../development/OCR_ONLINE_VALIDATION_LOG.md)。
- `--help` 末尾打印命名模板支持的变量清单（与设置页一致，来源 `ArchiveNamingRule.SupportedPlaceholders`）。

## 7. 自测与联调示例

```
:: 内置自测（可离线随时跑）
OfflineTester.exe --selftest

:: 拿到真实密钥后联调并保存脱敏返回
set BAIDU_OCR_AK=xxx
set BAIDU_OCR_SK=yyy
OfflineTester.exe sample --force-ocr --save-response ocr-raw

:: 用脱敏返回校准映射
OfflineTester.exe --parse-ocr-json ocr-raw\某发票.ocr.desensitized.json
```

## 8. 安全约束

- **不要提交样例 PDF**（`.gitignore` 已忽略 `sample/*.pdf`）。
- `--save-response` 保存的是**脱敏**版（长数字串/税号已掩码），但仍建议放入被忽略目录（如 `ocr-raw/`，`.gitignore` 已忽略），**不要提交**。
- 工具不落盘任何 AK/SK/token。

## 9. 样例提供与脱敏

用于校准的样例应满足：

- **脱敏但保留版式与字段标签**：可涂抹/替换姓名、手机号、身份证、纳税人识别号中的敏感位，但保留"发票号码/开票日期/价税合计/行程"等字段标签，否则无法校准解析。
- 尽量为**文本型 PDF**（能选中复制文字）；如有图片型/扫描件样例，单独提供 1 份用于验证 OCR 分支。
- 样例放 `sample/`（`.gitignore` 忽略），或运行时传入任意本地路径。

**记录/校准 OCR 返回**：`--save-response` 输出的脱敏 JSON 供 `--parse-ocr-json` 校准映射；真实返回 JSON（含姓名/号码/税号）必须脱敏后再共享，**不要提交**。校准 `BaiduInvoiceFieldMapper` 时对照 JSON 的 `type` 与 `result` 字段名更新映射表，并同步 [../requirements/FIELD_EXTRACTION_SPEC.md](../requirements/FIELD_EXTRACTION_SPEC.md)。

> **AK/SK/token 绝不写入代码、文档、日志、截图**；离线工具一律用环境变量传入。
