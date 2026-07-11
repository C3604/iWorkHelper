# 字段提取规范 FIELD_EXTRACTION_SPEC

> 日期：2026-07-06
> 对应模型：`Core/Invoice/InvoiceInfo.vb`、`InvoiceTripInfo.vb`、`InvoiceFieldNames.vb`。
> 用户要求“提取发票全部字段”，故模型尽量完整，并预留扩展字段。

## 1. 增值税电子发票字段

| 字段 | 模型属性 | 常量 (InvoiceFieldNames) | 必填/可选 | 命名参与 |
|------|---------|--------------------------|-----------|---------|
| 发票代码 | InvoiceCode | InvoiceCode | 可选 | 否 |
| 发票号码 | InvoiceNumber | InvoiceNumber | **必填** | 是 |
| 开票日期 | InvoiceDate | InvoiceDate | **必填** | 是（规整为 YYYYMMDD） |
| 购买方名称 | BuyerName | BuyerName | 可选 | 否 |
| 购买方纳税人识别号 | BuyerTaxId | BuyerTaxId | 可选 | 否 |
| 销售方名称 | SellerName | SellerName | 建议 | 是 |
| 销售方纳税人识别号 | SellerTaxId | SellerTaxId | 可选 | 否 |
| 项目名称 | ItemName | ItemName | 可选 | 否 |
| 规格型号 | Specification | Specification | 可选 | 否 |
| 单位 | Unit | Unit | 可选 | 否 |
| 数量 | Quantity | Quantity | 可选 | 否 |
| 单价 | UnitPrice | UnitPrice | 可选 | 否 |
| 金额 | Amount | Amount | 建议 | 否 |
| 税率 | TaxRate | TaxRate | 可选 | 否 |
| 税额 | TaxAmount | TaxAmount | 可选 | 否 |
| 价税合计 | TotalWithTax | TotalWithTax | **必填** | 是 |
| 备注 | Remark | Remark | 可选 | 否 |
| 校验码 | CheckCode | CheckCode | 可选 | 否 |
| 收款人 | Payee | Payee | 可选 | 否 |
| 复核人 | Reviewer | Reviewer | 可选 | 否 |
| 开票人 | Drawer | Drawer | 可选 | 否 |

> “必填”指用于判定识别是否“成功”的关键字段。
> **滴滴/行程单**：由 `LocalTextInvoiceRecognizer` 处理，命名核心字段为 乘车日期/金额/出发地点/到达地点。
> **常规发票（非滴滴）**：由 `GeneralInvoiceLocalRecognizer` 处理（2026-07-10 新增），命名核心字段为 **开票日期 + 金额 + 销售方名称**（`KeyFieldEvaluator.GetMissingGeneralInvoiceNamingFields`）；缺任一 → 部分成功 → 启用时回退在线 OCR。

## 1.1 常规发票本地识别（候选评分 + 分区 + 明细）

`GeneralInvoiceLocalRecognizer` 采用「逻辑行（坐标行重建/换行）+ 字段候选 + 评分择优 + 分区解析」，避免滴滴校准正则误用于常规发票。字段提取与评分要点（`GeneralInvoiceCandidateScorer`）：

| 字段 | 提取/评分策略 | 防误取 |
|------|--------------|--------|
| 开票日期 | 靠近「开票日期」标签 + 头部区 + 格式完整（`yyyy-MM-dd`/`/`/`.`/`年月日`）| 不取行程/商品日期 |
| 发票号码 | 靠近「发票号码」标签；长度 20（数电）/8（传统）最高，10–12 降分 | 不从「发票代码」行取；避开订单号/税号/校验码 |
| 发票代码 | 靠近「发票代码」标签，10/12 位 | 数电票允许为空 |
| 销售方名称 | 靠近「销售方/销方/销售方名称」标签或区域 + 组织特征词（有限公司/公司/经营部/店…）| 不被购买方/项目名称覆盖 |
| 购买方名称 | 靠近「购买方/购方」；与销售方按 X 坐标（有坐标）或行序（无坐标）分栏 | 购销不混淆 |
| 价税合计（{金额}）| **价税合计(小写) > 全局价税合计后首金额**；同行取「小写」侧最后一个两位小数 | 税额/税率降分；**不取明细行金额** |
| 税率 | `13/9/6/3/1/0%`、免税、不征税 | — |
| 税额 | 「合计…税额」区金额 | — |

`{金额}` 命名取值优先级（常规发票，`ArchiveNamingRule`）：**价税合计 > 金额**（行程金额为空）。

**数电票（全电发票）与长数字消歧策略**：
- 数电票通常**只有 20 位发票号码、无传统发票代码** → 发票代码允许为空（仅在有「发票代码」标签时才产候选）。
- 发票号码抽取只在含「发票号码」标签的行进行；提取的数字**后不得紧跟字母**（`(?![0-9A-Za-z])`），避免取到「91310000MA…」这类统一社会信用代码/纳税人识别号的数字前缀。
- 号码所在行若含「纳税人识别号 / 统一社会信用代码 / 校验码 / 订单号 / 流水号」→ `LongNumberContextPenalty` 降分，确保靠近「发票号码」标签的号码优先，不误取税号/校验码/订单号。
- 冲突以**候选评分**择优，不用单条正则硬匹配。

## 1.2 商品明细字段（InvoiceLineItem）

`InvoiceInfo.LineItems`（2026-07-10 新增），由 `GeneralInvoiceLineItemParser` 在明细表格区（`PdfTableRegionDetector`：表头「项目名称/货物+金额/税额」→ 到「合计/价税合计」之间）逐行尽力解析。

| 字段 | 属性 | 说明 |
|------|------|------|
| 项目名称 | ItemName | 优先 `*类别*名称`；无 `*` 且几乎无数字的行视为名称换行续行 |
| 规格型号/单位/数量/单价 | Specification/Unit/Quantity/UnitPrice | 允许为空 |
| 金额/税率/税额 | Amount/TaxRate/TaxAmount | 行内多个两位小数：末位=税额、其前=金额；仅一个=金额 |
| 原始行文本/行号/页码 | RawLine/LineIndex/PageNumber | 诊断用 |

规则：支持一行/多行/名称换行/简化明细；**单行失败跳过，明细解析失败不影响头部字段**；明细金额+税额与价税合计不一致仅记「提示」，不判失败。

## 2. 网约车/滴滴行程单字段

模型：`InvoiceTripInfo`（一张行程单含多条 `Trips`）。

| 字段 | 属性 | 常量 | 说明 |
|------|------|------|------|
| 乘车人 | Passenger | Passenger | 表头级，常在行程单顶部 |
| 出发时间 | DepartureTime | DepartureTime | 每条行程 |
| 到达时间 | ArrivalTime | ArrivalTime | 每条行程 |
| 起点 | StartLocation | StartLocation | 每条行程 |
| 终点 | EndLocation | EndLocation | 每条行程 |
| 车型/服务类型 | ServiceType | ServiceType | 快车/专车等 |
| 订单号 | OrderNumber | OrderNumber | 每条行程 |
| 行程金额 | TripAmount | TripAmount | 每条行程 |
| 城市 | City | City | 表头或每条 |
| 司机/车辆信息 | DriverInfo | DriverInfo | 如可识别 |

> 当前 `LocalTextInvoiceRecognizer.ParseTripFields` 仅解析**表头级**字段（乘车人、城市、合计），**逐条行程明细待样例校准**。

## 3. 未知/扩展字段策略

- `InvoiceInfo.ExtendedFields`（`Dictionary(Of String, String)`）承载模型未显式定义的字段。
- 用途：百度 OCR 或未来其它来源返回的额外字段，一律进此字典，**确保不丢字段**。
- 写入：`InvoiceInfo.SetExtendedField(name, value)`。

## 4. 值的表示约定

- 金额/数量/日期等**一律以字符串保存原始文本**，模型层不做强制数值/日期转换。
- 原因：票据版式差异大，过早转换易失败或丢精度；命名环节按需做轻量规整（如日期取 8 位数字）。

## 5. 命名规则中的字段使用

默认文件名：`开票日期_销售方名称_发票号码_价税合计.pdf`（`ArchiveNamingRule`）。
- 开票日期规整为 `YYYYMMDD`（提取其中数字前 8 位）。
- 任一字段缺失则跳过该片段；全部缺失时回退为 `原附件名` 或 `未识别票据_时间戳`。
- 所有片段经 `FileNameSanitizer` 清理 Windows 非法字符。

## 6. 本地正则规则现状（待校准）

`LocalTextInvoiceRecognizer` 当前正则为**初版**，示例：
- 发票号码：`发票号码[:：\s]*([0-9]{8,20})`
- 价税合计：`价税合计[（(]?小写[）)]?[:：\s]*[￥¥]?\s*([0-9]+\.?[0-9]*)`
- 销售方名称：`销售方[\s\S]{0,10}?名\s*称[:：\s]*([^\r\n]{2,40})`

**校准计划**：拿到 2~3 份脱敏样例后：
1. 核对每条正则命中率，修正边界（购/销方名称易串行、金额易漏 ￥）。
2. 补充行程单逐条明细的解析。
3. 记录 PdfPig 抽取文本的实际排版特征（是否保留换行、字段是否同行）。

## 6.1 真实样例校准结果（第三阶段，5 份脱敏滴滴合并 PDF）

用离线工具对 `/sample` 下 5 份 PDF 实测（PdfPig 抽取 + 本地解析），结论：

- 均为**文本型 PDF**（2 页：第 1 页电子发票 + 第 2 页行程单），本地解析全部成功。
- 文本层无空格分隔、字段紧邻（如金额与税率拼接 `134.433%`），正则需按位置切分。
- 可靠字段（本地正则命中率 100%）：发票号码、开票日期、销售方名称（=滴滴出行科技有限公司）、购买方名称、价税合计、行程日期、行程金额、车型。
- 已知局限：
  - **税率**在拼接文本中仅能取个位（滴滴为 3%，正确；13%/16% 类两位税率需 OCR 或进一步切分）。
  - **项目名称**偶有截断（不影响命名与关键字段）。
  - 滴滴新版电子发票**无独立发票代码**（仅 20 位发票号码）。
- 购/销方判定：名称含"滴滴"者判为销售方，另一为购买方；无滴滴标识时按位置回退。

## 6.2 用户实际命名习惯（观察，供后续可选校准）

用户提供的样例文件名形如 `2026-05-18_138.46_滴滴出行_合并.pdf`，即
`行程日期_价税合计_滴滴出行_合并`。与本项目**默认命名规则**（`开票日期_销售方_发票号码_价税合计`）不同：
- 用户用**行程日期**（2026-05-18）而非开票日期（2026-05-26）；
- 用**滴滴出行**简称而非全称；含"合并"后缀。

当前按需求既定默认规则实现。如需贴合用户习惯，可作为**可选后续项**：增加"命名模板可配置"或滴滴专用命名分支。

## 7. 百度 OCR 字段映射表（初版，待真实返回校准）

映射集中在 `Core/Ocr/Baidu/BaiduInvoiceFieldMapper.vb`。增值税发票部分（百度公开字段名 → 内部字段）：

| 百度返回字段 | 映射到 | 备注 |
|-------------|--------|------|
| InvoiceNum | 发票号码 | |
| InvoiceCode | 发票代码 | 新版电子票可能为空 |
| InvoiceDate | 开票日期 | |
| SellerName | 销售方名称 | |
| SellerRegisterNum | 销售方纳税人识别号 | |
| PurchaserName | 购买方名称 | |
| PurchaserRegisterNum | 购买方纳税人识别号 | |
| CommodityName | 项目名称 | 数组，多行 |
| CommodityType | 规格型号 | |
| CommodityUnit | 单位 | |
| CommodityNum | 数量 | |
| CommodityPrice | 单价 | |
| CommodityTaxRate | 税率 | |
| TotalAmount | 金额 | 合计不含税 |
| TotalTax | 税额 | 合计税额 |
| AmountInFiguers / AmountInFigures | 价税合计 | 百度实际拼写为 AmountInFiguers |
| Checkcode / CheckCode | 校验码 | |
| NoteDrawer | 开票人 | |
| Payee | 收款人 | |
| Reviewer | 复核人 | |
| Remarks | 备注 | |

> ⚠️ 上表为初版映射（基于百度公开文档字段名）；第 8 节为**基于真实返回校准后**的结果。样例脱敏与校准方法见 [../testing/OFFLINE_TESTER_GUIDE.md](../testing/OFFLINE_TESTER_GUIDE.md)。

### 未知字段策略
- **未在映射表中的百度字段一律写入 `InvoiceInfo.ExtendedFields`（保留百度原始字段名）**，确保不丢失。
- 每个字段（含未知）同时登记到 `InvoiceRecognitionResult.Fields`，`Source="BaiduOcr:原始字段名"`，并保留置信度。

### 网约车行程单（taxi_online_ticket / taxi_receipt）策略
- 百度返回字段名不稳定，当前实现：**全部字段入 ExtendedFields**，并按字段名关键字（total/amount/start/end/time/city/order/type/passenger/mile）**尽力回填**到 `InvoiceTripInfo` 与发票级金额/日期。
- **按 RowIndex 分组支持多条行程**（若百度以数组返回多行，每个 RowIndex 生成一条 `InvoiceTripInfo`）。
- 真实字段名以样例返回校准后再细化强类型映射。

## 8. 规范字段映射表（第六阶段，已基于真实返回校准）

> **验证状态说明（第六阶段更新）**：已用真实 AK/SK 对 5 份真实滴滴合并 PDF 完成 `multiple_invoice` 真实联调，下列增值税发票字段名**已通过真实 OCR 返回验证**（字段名，值不记录）。行程字段名亦来自真实返回（`Passeng*`）。合成 JSON 仅用于回归自测。
>
> 验证状态取值：`真实OCR`（已通过真实返回验证）/ `本地样例`（仅本地文本样例验证）/ `合成`（仅合成 JSON 验证）/ `待验证`。

### 8.1 增值税发票字段（真实返回已验证）
| 内部字段 | 百度原始字段 | 关键 | 参与命名 | 可空 | 验证状态 | 备注 |
|---------|-------------|------|---------|------|---------|------|
| 发票号码 | InvoiceNum | 是 | 是 | 否 | 真实OCR | 20 位 |
| 发票代码 | InvoiceCode | 否 | 否 | 是 | 真实OCR | 新版电子票常空 |
| 开票日期 | InvoiceDate | 是 | 是 | 否 | 真实OCR | 规整 YYYYMMDD |
| 销售方名称 | SellerName | 是 | 是 | 是 | 真实OCR | =滴滴出行科技有限公司 |
| 销售方纳税人识别号 | SellerRegisterNum | 否 | 否 | 是 | 真实OCR | |
| 购买方名称 | PurchaserName | 否 | 否 | 是 | 真实OCR | |
| 购买方纳税人识别号 | PurchaserRegisterNum | 否 | 否 | 是 | 真实OCR | |
| 项目名称 | CommodityName | 否 | 否 | 是 | 真实OCR | 带 row，多行 |
| 税率 | CommodityTaxRate | 否 | 否 | 是 | 真实OCR | 带 row |
| 金额 | TotalAmount | 是(替代) | 否 | 是 | 真实OCR | 合计不含税 |
| 税额 | TotalTax | 否 | 否 | 是 | 真实OCR | |
| 价税合计 | AmountInFiguers | 是 | 是 | 是 | 真实OCR | 百度拼写 Figuers |
| 开票人 | NoteDrawer | 否 | 可入模板 | 是 | 真实OCR | |
| 复核人 | **Checker** | 否 | 否 | 是 | 真实OCR | **真实字段名为 Checker（非 Reviewer）** |
| 备注 | Remarks | 否 | 否 | 是 | 真实OCR | |
| 收款人 | Payee | 否 | 否 | 是 | 待验证 | 本次样例为空 |
| 校验码 | CheckCode/Checkcode | 否 | 否 | 是 | 待验证 | 本次样例为空 |
| 规格型号 | CommodityType | 否 | 否 | 是 | 待验证 | 本次样例为空 |
| 单位 | CommodityUnit | 否 | 否 | 是 | 待验证 | 本次样例为空 |
| 数量 | CommodityNum | 否 | 否 | 是 | 真实OCR | 带 row |
| 单价 | CommodityPrice | 否 | 否 | 是 | 真实OCR | 带 row |
| 票据类型 | (item.type) | 是 | 可入模板 | 否 | 真实OCR | 见类型映射 |

### 8.2 滴滴旅客运输发票内嵌行程字段（真实返回已验证）
> 重要：滴滴合并发票 OCR 为 `vat_invoice` 时，行程信息**内嵌**在 result 的 `Passeng*` 字段（带 `row`），而非独立 `taxi_online_ticket` 项。

| 内部字段 | 百度原始字段 | 参与命名 | 可空 | 验证状态 | 备注 |
|---------|-------------|---------|------|---------|------|
| 乘车人 | PassengName | 是 | 是 | 真实OCR | |
| 出发日期 | PassengDate | 是 | 是 | 真实OCR | 亦写入 ExtendedFields[行程起止日期] |
| 起点 | PassengOrigin | 是 | 是 | 真实OCR | OCR 起终点比本地拆分更干净 |
| 终点 | PassengDestination | 是 | 是 | 真实OCR | |
| 车型/服务类型 | PassengVehicleType | 否 | 是 | 真实OCR | |
| （身份证/等级等） | PassengIdNum/PassengClass | 否 | 是 | 真实OCR | 入 ExtendedFields（敏感，不参与命名） |

### 8.3 未知字段与其它
- 未在上表的返回字段（`AmountInWords / InvoiceType / InvoiceTypeOrg / SellerBank / SellerAddress / PurchaserBank / Transport* / City / Province ...`）→ 一律进 `ExtendedFields`（保留百度原始字段名），验证状态 `真实OCR`（确认存在但未强类型化）。
- `taxi_online_ticket / taxi_receipt` 独立行程单类型的真实字段名：`待验证`（本次样例均为 vat_invoice 内嵌行程，未出现独立行程单类型）。

### 8.4 独立行程单类型兼容（第七阶段，兼容性预留）
- `BaiduInvoiceTypeMapper`：`taxi_online_ticket`→行程单、`taxi_receipt`→行程单、未知 type→Other（不失败）。
- `BaiduInvoiceFieldMapper.AssignRideField`：按字段名关键词（origin/start/起、dest/终、arriv/到达、vehicle/车型、mile/里程、order/订单、city/城市、passeng/乘车人、time/date/depart/时间、total/amount/fare/金额）尽量提取乘车人/出发时间/到达时间/起点/终点/城市/车型/订单号/里程/行程金额，按 `row` 分组。
- 验证状态：`合成`（`--selftest` 通过）+ `兼容性预留`——**真实独立行程单返回未获取**（本次样例均为 vat_invoice 内嵌 `Passeng*`）。真实字段名待样例校准。
- 命名仍按现有行程单模板；当前一单一条行程，不拆分多条，不处理图片附件。

| 内部字段 | 百度原始字段 | 票据类型 | 关键 | 参与命名 | 可空 | 备注 |
|---------|-------------|---------|------|---------|------|------|
| 发票号码 | InvoiceNum | 增值税发票 | 是 | 是 | 否 | 20 位 |
| 发票代码 | InvoiceCode | 增值税发票 | 否 | 否 | 是 | 新版电子票常空 |
| 开票日期 | InvoiceDate | 增值税发票 | 是 | 是 | 否 | 规整 YYYYMMDD |
| 销售方名称 | SellerName | 增值税发票 | 是 | 是 | 是 | 滴滴场景=滴滴出行科技有限公司 |
| 销售方纳税人识别号 | SellerRegisterNum | 增值税发票 | 否 | 否 | 是 | |
| 购买方名称 | PurchaserName | 增值税发票 | 否 | 否 | 是 | |
| 购买方纳税人识别号 | PurchaserRegisterNum | 增值税发票 | 否 | 否 | 是 | |
| 项目名称 | CommodityName | 增值税发票 | 否 | 否 | 是 | 数组多行 |
| 规格型号 | CommodityType | 增值税发票 | 否 | 否 | 是 | |
| 单位 | CommodityUnit | 增值税发票 | 否 | 否 | 是 | |
| 数量 | CommodityNum | 增值税发票 | 否 | 否 | 是 | |
| 单价 | CommodityPrice | 增值税发票 | 否 | 否 | 是 | |
| 税率 | CommodityTaxRate | 增值税发票 | 否 | 否 | 是 | |
| 金额 | TotalAmount | 增值税发票 | 是(替代) | 否 | 是 | 价税合计缺失时替代 |
| 税额 | TotalTax | 增值税发票 | 否 | 否 | 是 | |
| 价税合计 | AmountInFiguers | 增值税发票 | 是 | 是 | 是 | 百度拼写 Figuers |
| 校验码 | Checkcode/CheckCode | 增值税发票 | 否 | 否 | 是 | |
| 开票人 | NoteDrawer | 增值税发票 | 否 | 否(可入模板) | 是 | |
| 收款人 | Payee | 增值税发票 | 否 | 否 | 是 | |
| 复核人 | Reviewer | 增值税发票 | 否 | 否 | 是 | |
| 备注 | Remarks | 增值税发票 | 否 | 否 | 是 | |
| 票据类型 | (type 字段) | 通用 | 是 | 否 | 否 | 见类型映射 |
| 乘车人 | (待校准) | 行程单 | 否 | 是 | 是 | 样例常脱敏为空 |
| 出发时间 | (待校准) | 行程单 | 否 | 是 | 是 | 首条行程 |
| 到达时间 | (待校准) | 行程单 | 否 | 否 | 是 | |
| 起点 | (待校准) | 行程单 | 否 | 是 | 是 | |
| 终点 | (待校准) | 行程单 | 否 | 是 | 是 | |
| 城市 | (待校准) | 行程单 | 否 | 否 | 是 | |
| 车型/服务类型 | (待校准) | 行程单 | 否 | 否 | 是 | |
| 订单号 | (待校准) | 行程单 | 否 | 是 | 是 | |
| 行程金额 | (待校准) | 行程单 | 否 | 是 | 是 | |
| 里程 | (待校准) | 行程单 | 否 | 否 | 是 | |
| 未知字段 | 原样保留 | 任意 | 否 | 否 | 是 | 一律入 ExtendedFields |

## 8.5 本地行程单解析（坐标行重建，本地识别稳定性专项）

- **本地起点/终点/金额改用坐标按列取单元格**（`PdfTextLayoutExtractor` + `LocalTextInvoiceRecognizer.ParseTripTableByLayout`），不再依赖 `page.Text` 全文正则切分：
  - 读取行程单表头列 X（每份 PDF 列宽自适应），数据词按最近列中心归入起点/终点；
  - **列区间过滤**排除车型/城市/时间/里程/金额/备注列（含未知车型别名如“滴滴特快”）；
  - **同列跨行词按 Y→X 拼接**，支持长地址换行；
  - 文本先经 `LocalTextNormalizer` 归一化（全/半角、冒号、空格、金额符号、多格式日期）。
- 命名核心字段（乘车日期/金额/出发地点/到达地点）由 `KeyFieldEvaluator.IsNamingSufficient` 判定；本地不足回退在线 OCR。
- 验证状态：`本地样例`（5 份真实样例命中 5/5，见 LOCAL_TEXT_RECOGNITION_DIAGNOSTIC_REPORT）。图片型/异版式仍走 OCR。

## 9. 多条行程明细（第四阶段）

- 模型：`InvoiceInfo.Trips`（`List(Of InvoiceTripInfo)`）+ `InvoiceInfo.StatedTripCount`（"共N笔行程"）。
- `InvoiceTripInfo` 字段：乘车人、出发时间、到达时间、起点、终点、城市、车型/服务类型、订单号、行程金额、里程、序号。
- **本地解析**（已用真实样例验证）：以"上车时间（MM-DD HH:MM）"为锚点逐行提取；起点/终点尽力拆分（近似）。单行失败不影响其它行与整份 PDF。
- **百度解析**：按 RowIndex 分组构建多条；优先使用百度结构化明细。
- **命名**：默认取**首条行程**核心信息；归档结果记录 `TripCount`。
- **归档单位仍为 PDF**（多条行程不拆分为多个文件）。
- 现有 5 份样例均为**单条行程**，多条场景的实测待真实多行程样例。
