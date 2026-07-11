# 内网版 / 外网版编译设计 BUILD_VARIANTS

> 更新日期：2026-07-11
> 实现：`iWorkhelper.vbproj`（配置定义）、`Core/Common/BuildFeatures.vb`（编译期开关）。

## 1. 四套编译配置

| 配置 | 条件编译常量 | 输出目录 | 在线 OCR | 优化 |
|------|-------------|---------|---------|------|
| Debug | （无 INTERNET_BUILD） | `bin\Debug\` | **禁用** | 否 |
| Release | （无 INTERNET_BUILD） | `bin\Release\` | **禁用** | 是 |
| **Release-Intranet**（内网版） | `INTRANET_BUILD` | `bin\Release-Intranet\` | **禁用** | 是 |
| **Release-Internet**（外网版） | `INTERNET_BUILD` | `bin\Release-Internet\` | **启用** | 是 |

依据 `iWorkhelper.vbproj` 中四个 `PropertyGroup Condition` 段。

## 2. 功能差异开关（BuildFeatures）

`Core/Common/BuildFeatures.vb` 是编译期功能开关的**单一来源**，按条件编译常量选择：

```
#If INTERNET_BUILD Then     → OnlineParserEnabled = True，  EditionDisplay = ""
#Else                       → OnlineParserEnabled = False， EditionDisplay = "（内网版）"
```

**默认安全策略**：未定义 `INTERNET_BUILD` 即视为内网版，**在线解析被禁用**。因此 Debug、Release、Release-Intranet 均禁用在线 OCR；仅 Release-Internet 启用。

## 3. 开关的实际作用点

`BuildFeatures.OnlineParserEnabled` 真正门控以下行为（源码引用）：

- `Core/Recognition/BaiduOcrInvoiceRecognizer.vb`：`OnlineParserEnabled=False` 时**直接拒绝在线 OCR**，返回配置缺失，仅走本地识别。
- `SettingsForm.vb`：内网版**隐藏** OCR 配置分组（`grpOcr`）与"在线解析"单选项；设置界面版本号后追加 `EditionDisplay`（"（内网版）"）。

因此：
- **内网版**：设置页无 OCR 配置项，归档只用本地文本识别，不产生任何外网请求（符合内网合规要求）。
- **外网版**：设置页显示 OCR 配置，本地不足时可回退百度 OCR。

## 4. 构建命令

```
MSBuild iWorkhelper.sln /p:Configuration=Debug           /p:Platform=AnyCPU
MSBuild iWorkhelper.sln /p:Configuration=Release-Intranet /p:Platform=AnyCPU
MSBuild iWorkhelper.sln /p:Configuration=Release-Internet /p:Platform=AnyCPU
```

## 5. 维护约束

- **不可违背的约束**：内网版禁止任何在线 OCR/外网请求。修改识别或配置逻辑时，必须保持 `BuildFeatures.OnlineParserEnabled=False` 分支下无外网调用。
- 新增涉及外网的功能时，应统一通过 `BuildFeatures.OnlineParserEnabled` 门控，不要绕过该开关直接调用网络。
- 签名：发布使用 `ManifestKeyFile`（临时证书 `iWorkhelper_TemporaryKey.pfx`）+ `ManifestCertificateThumbprint`。真实私钥文件不入库，正式发布需配置正式代码签名证书。
