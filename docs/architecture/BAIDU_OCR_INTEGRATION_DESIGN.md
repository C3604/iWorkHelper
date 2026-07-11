# 百度 OCR 接入设计 BAIDU_OCR_INTEGRATION_DESIGN

> 日期：2026-07-06（第三阶段更新为**真实实现方案**）
> 接口：百度智能云 “智能财务票据识别”（multiple_invoice）。
> 实现代码：`Core/Ocr/Baidu/*`、`Core/Recognition/BaiduOcrInvoiceRecognizer.vb`、`Core/Recognition/RecognitionPipeline.vb`。

## 1. 组件总览

| 组件 | 职责 |
|------|------|
| `BaiduOcrOptions` | 配置 POCO（AK/SK/接口/超时/开关等），解耦 My.Settings |
| `OcrConfigProvider` | 从 My.Settings（主）/ XML（兼容）加载配置 |
| `BaiduAccessTokenProvider` | 换取 + 缓存 access_token |
| `BaiduOcrHttpClient` | 调用 multiple_invoice 接口 |
| `BaiduMultipleInvoiceResponseParser` | JavaScriptSerializer 解析返回 |
| `BaiduInvoiceTypeMapper` | 票据类型映射 |
| `BaiduInvoiceFieldMapper` | 字段集中映射 |
| `BaiduOcrInvoiceRecognizer` | 编排上述组件，产出识别结果 |
| `RecognitionPipeline` | 本地/在线调度与来源标记 |

## 2. 配置项（SettingsForm 为主入口）

设置界面（`SettingsForm`）已提供以下项，写入 `My.Settings`：

| UI 项 | 设置键 | 默认 |
|-------|--------|------|
| 启用在线 OCR | OcrEnabled | False |
| API Key | BaiduApiKey | 空 |
| Secret Key（密码遮罩） | BaiduSecretKey | 空 |
| 接口地址 | BaiduOcrApiUrl | https://aip.baidubce.com/rest/2.0/ocr/v1/multiple_invoice |
| Token 地址 | BaiduTokenUrl | https://aip.baidubce.com/oauth/2.0/token |
| 超时(毫秒) | OcrTimeoutMs | 30000 |
| 返回置信度 probability | OcrReturnProbability | True |
| 返回位置 location | OcrReturnLocation | False |
| 验真 verify_parameter | OcrVerifyParameter | False |
| 最大识别页数 | OcrMaxPages | 1 |
| 优先本地文本解析 | PreferLocalParse | True |
| 本地不足自动在线 OCR | AutoFallbackToOcr | True |

- **不硬编码密钥**；日志仅打印脱敏摘要（`MaskSecret`：前 2 后 2，中间 *）。
- **（第四阶段）Secret Key 经 DPAPI 加密存储**：`My.Settings.BaiduSecretKey` 存 `DPAPI:` 前缀密文，由 `ProtectedSettingsProvider` 透明加解密，`OcrConfigProvider` 解密后注入。详见 [../security/SECRET_STORAGE_DESIGN.md](../security/SECRET_STORAGE_DESIGN.md)。
- **（第七阶段）加密配置文件落地**：`%AppData%\iWorkHelper\baidu-ocr.config.xml`（`BaiduXmlConfigStore`）——SK 以 DPAPI 加密保存，`OcrConfigProvider` 在 My.Settings 未填 AK/SK 时**完整读取该文件并解密**（含 Enabled/URL/开关）。离线工具 `--save-baidu-config`/`--use-config` 写入/读取，已实测“加密落地→解密→真实 OCR”闭环。`.gitignore` 已忽略该文件，不入库。

## 3. Access Token（`BaiduAccessTokenProvider`）

- 请求：`POST {TokenUrl}?grant_type=client_credentials&client_id={AK}&client_secret={SK}`。
- 解析：`JavaScriptSerializer` → `access_token` / `expires_in`。
- **缓存**：进程内 `Shared`，按 AK 区分；`过期时刻 = UtcNow + (expires_in - 300s)`（提前 5 分钟刷新）。
- 不逐 PDF 重取；保存设置（密钥可能变更）时 `InvalidateCache()`。
- 失败返回结构化 `BaiduAccessTokenResult`（含 HTTP 码、error），不抛 UI；**token 不写日志**。

## 4. 识别请求（`BaiduOcrHttpClient`）

- `POST {OcrApiUrl}?access_token=xxx`，`Content-Type: application/x-www-form-urlencoded`。
- Body：`pdf_file`（PDF 字节 → Base64 → **`WebUtility.UrlEncode`**，不能用 `Uri.EscapeDataString`，后者对超长字符串有 ~65520 上限会抛异常）+ `pdf_file_num`（页码）+ `probability`/`location`/`verify_parameter`（按配置 true/false）。
- **不使用 image / url**；不处理图片附件。
- 超时 `TimeoutMilliseconds`；**TLS 1.2** 显式启用（`ServicePointManager.SecurityProtocol |= Tls12`）。
- 返回封装 `BaiduOcrRawResponse`（HTTP 码、原始 JSON、error_code/error_msg）。

## 5. 解析与字段映射

### 解析（`BaiduMultipleInvoiceResponseParser`）
- `JavaScriptSerializer.DeserializeObject` 弱类型；读取 `words_result_num` / `pdf_file_size` / `words_result[]`。
- 每项含 `type` + `result`；`result` 字段值泛化处理：字符串 / `{word,probability,location}` / 数组（多行）。
- 保留原始字段名、值、置信度（probability.average）、位置（location）。

### 映射（`BaiduInvoiceFieldMapper`，集中）
- 增值税发票字段映射表（百度名 → 内部名），示例：
  `InvoiceNum→发票号码`、`InvoiceDate→开票日期`、`SellerName→销售方名称`、`AmountInFiguers→价税合计`、`TotalAmount→金额`、`TotalTax→税额`、`CommodityTaxRate→税率` 等。
- **未知字段一律进 `ExtendedFields`（保留百度原始字段名）**，不丢失。
- 置信度写入 `InvoiceField.Confidence`；`Source="BaiduOcr:原始字段名"` 保留来源。
- 出行类（taxi_online_ticket/taxi_receipt）：字段名不稳定，全部入扩展字段并尽力回填金额/日期/起终点到 trip。
- **（第四阶段）按 RowIndex 分组支持多条行程**：百度以数组返回多行时，每个 RowIndex 生成一条 `InvoiceTripInfo`；优先使用百度结构化明细。
- 单字段失败 `Try/Catch` 吞掉，不影响整体。
- 空字段：接口成功但无字段 → `PartialSuccess`（"识别成功但未提取到字段"），非通用失败。

## 6. 调用策略（`RecognitionPipeline`）

```
本地抽取 + 本地解析（始终执行，成本低，且作为 OCR 失败兜底）
  ├ PreferLocalParse 且本地关键字段充分 → 采用本地（来源 LocalText），不调 OCR
  └ 否则：
       canOnline = OcrEnabled 且 AK/SK/URL 齐全
       wantOnline = 非本地优先 或 AutoFallbackToOcr 或 疑似图片型
       若 canOnline 且 wantOnline → 调用百度 OCR
            ├ OCR 可用 → 采用 OCR 字段（本地也有字段→Mixed；否则 BaiduOcr）
            ├ OCR 失败但本地有部分字段 → 本地部分成功（日志记录原因）
            └ 都失败 → Failed / ConfigurationMissing
       否则 → 本地可用则用本地；本地需 OCR 但不可用 → NeedsOcr（附原因）
```

关键字段（`KeyFieldEvaluator`）：发票号码、开票日期、销售方名称、价税合计或金额、票据类型。

## 7. 错误处理约定

| 情况 | 状态 |
|------|------|
| 未启用 | ConfigurationMissing（"在线 OCR 未启用"） |
| AK/SK/URL 不全 | ConfigurationMissing（"配置缺失"） |
| 文件过大（>约 6MB 原始） | Failure（"文件过大，无法在线 OCR"） |
| Token 获取失败 | Failure（含 HTTP 码/error，脱敏） |
| error_code / error_msg | 该页失败并记录，多页时继续；全失败→Failure |
| 成功但无字段 | PartialSuccess（"识别成功但未提取到字段"） |
| 成功且字段齐 | Success |

## 7.1 验证状态（第六阶段：真实联调成功）

- **真实在线联调：已成功**（真实 AK/SK，会话级环境变量，未入库）。5 份样例 `Success/Mixed/VatInvoice/31 字段`。
- **联调修复的真实 bug**：`pdf_file` 编码原用 `Uri.EscapeDataString`（.NET Framework 约 65520 字符上限）对超长 Base64 抛异常 → 改 `System.Net.WebUtility.UrlEncode`（无长度限制）。
- **真实返回校准**：
  - 复核人真实字段名 `Checker`（补入映射）。
  - 滴滴发票行程信息内嵌于 `vat_invoice.result` 的 `Passeng*` 字段（带 `row`）→ VAT 映射据 `row` 分组构建行程明细。
  - 解析器读取明细元素 `row` 作为行号。
- **安全**：日志无 AK/SK/token 明文（已搜索核验）；脱敏返回仅存仓库外目录。
- 详见 [../development/OCR_ONLINE_VALIDATION_LOG.md](../development/OCR_ONLINE_VALIDATION_LOG.md) 与 [../requirements/FIELD_EXTRACTION_SPEC.md](../requirements/FIELD_EXTRACTION_SPEC.md) 第 8 节。

## 8. 后续待实现（TODO）

1. **真实 AK/SK 联调**：用离线工具 `--ocr` 抓取真实返回 JSON，校准字段映射表（第 5 节）。
2. **taxi_online_ticket 字段名核对**：以真实返回确认，替换"尽力猜测"。
3. **多条行程明细**：拆分逐条起点/终点/时间。
4. **SK 加密存储**：user.config 明文 → DPAPI（建议）。
5. 大文件/图片型 PDF 若确需支持：评估 PDF→图片渲染（当前仅 pdf_file 文本/矢量）。
