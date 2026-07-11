# 百度 OCR 在线联调记录 OCR_ONLINE_VALIDATION_LOG

> 日期：2026-07-07
> 状态：**已使用真实 AK/SK 完成真实在线联调（multiple_invoice 智能财务票据识别）。**
> 安全：AK/SK 仅以会话级环境变量在运行时使用，**未写入任何代码/文档/提交文件**；日志已核验无 AK/SK/token 明文；脱敏返回仅存于仓库外临时目录，未入库。

## 1. 结论

- **multiple_invoice 接口：真实联调成功。** 5 份真实滴滴合并 PDF 全部 `状态=Success，来源=Mixed，类型=VatInvoice，字段数=31`。
- 联调中发现并修复 **1 个真实 bug**（详见第 3 节）。
- 基于真实返回**校准了字段映射**（详见 FIELD_EXTRACTION_SPEC 第 8 节 与 第 4 节）。

## 2. 联调验证项（真实执行结果）

| # | 验证项 | 结果 | 证据 |
|---|--------|------|------|
| 1 | Access Token 获取成功 | **通过** | OCR 成功返回 HTTP 200；错误密钥时返回 401 |
| 2 | Token 缓存生效 | **通过** | 5 份 PDF 在同一进程内仅取 1 次 token（日志累计次数=独立进程数） |
| 3 | 多 PDF 不重复取 token | **通过** | 同上（进程内 Shared 缓存） |
| 4 | `pdf_file` 参数可用 | **通过** | PDF 二进制 Base64+UrlEncode 提交，接口正常识别 |
| 5 | `pdf_file_num` 可用 | **通过** | 第 1 页识别成功（HTTP 200，成功页=1） |
| 6 | multiple_invoice 返回可解析 | **通过** | 弱类型解析出 31 字段并映射 |
| 7 | 密钥错误明确报错 | **通过** | “Token 获取失败：Client authentication failed（HTTP 401）”，并回退本地 |
| 8 | 未授权/额度不足明确报错 | **部分**（错误处理路径已验，未专门触发额度不足） | 错误码/错误信息保留逻辑就绪 |
| 9 | 超时明确报错 | **未专门触发** | 代码有 Timeout + 网络异常捕获 |
| 10 | 字段为空返回“识别成功但字段为空” | **未触发**（真实返回均有字段） | 逻辑就绪 |
| 11 | 单 PDF 失败不影响批量 | **通过** | 错误密钥时该项回退本地，不中断 |
| 12 | 日志无 AK/SK/token 明文 | **通过** | 对日志搜索真实 AK/SK/`access_token=` 均为 False |
| 13 | 脱敏响应无敏感信息 | **通过** | 长数字串（税号/身份证/手机号/发票号）已掩码 |
| 14 | 仅处理 PDF，不新增图片路径 | **通过** | 仅 `pdf_file`，未引入图片识别 |

## 3. 联调发现并修复的真实 Bug

- **现象**：首次真实调用报 “无效的 URI: URI 字符串太长（HTTP 0）”，OCR 全页失败并回退本地。
- **根因**：`BaiduOcrHttpClient` 用 `Uri.EscapeDataString(base64)` 对 PDF 的 Base64（约 22 万字符）做 URL 编码；**.NET Framework 的 `Uri.EscapeDataString` 有约 65520 字符上限**，超限抛异常。
- **修复**：改用 `System.Net.WebUtility.UrlEncode`（无长度限制），正确编码 `+ / =`。
- **修复后**：5 份样例全部 OCR 成功，字段数 31。

## 4. 基于真实返回的字段映射校准

真实 `vat_invoice` 返回结构（字段名，值不记录）：
- `result` 内每个字段为 `[{probability:{average,min}, word, row?}]` 结构；`type` 与 `result` 同级位于 words_result 项内。
- **校准点**：
  1. **复核人真实字段名为 `Checker`**（此前误用 `Reviewer`）→ 已补 `Checker`→复核人。
  2. **滴滴旅客运输发票的行程信息内嵌在 `vat_invoice.result`** 的 `Passeng*` 字段（`PassengName/PassengDate/PassengOrigin/PassengDestination/PassengVehicleType`，带 `row`）→ 已在 VAT 映射中据 `row` 分组构建行程明细；OCR 路径行程明细数量由 0 修正为 1。
  3. 明细字段元素带 `"row"` → 解析器已读取 `row` 作为行号，支持多条明细。
  4. 已确认可映射字段名：`InvoiceNum / InvoiceDate / SellerName / SellerRegisterNum / PurchaserName / PurchaserRegisterNum / CommodityName / CommodityTaxRate / AmountInFiguers / TotalAmount / TotalTax / NoteDrawer / Remarks`。
  5. 其它返回字段（`AmountInWords / InvoiceType / InvoiceTypeOrg / Seller/PurchaserBank/Address / Passeng*(其余) / Transport*` 等）→ 进 ExtendedFields，不丢失。

## 5. 用户复现步骤（本机手工提供 AK/SK）

```
set BAIDU_OCR_AK=你的APIKey
set BAIDU_OCR_SK=你的SecretKey
OfflineTester.exe sample --force-ocr --save-response ocr-raw
OfflineTester.exe --parse-ocr-json ocr-raw\某发票.ocr.desensitized.json
```
- `ocr-raw/` 已被 `.gitignore` 忽略；脱敏返回也不要提交。

## 6. 待真实联调补充（TODO）
- 主动触发并记录：接口超时、额度不足/未授权、真实空字段返回。
- 多页 PDF（MaxPages>1）逐页真实验证。
- 多条行程真实样例（现有均单条）。
