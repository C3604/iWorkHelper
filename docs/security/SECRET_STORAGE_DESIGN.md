# 敏感配置存储设计 SECRET_STORAGE_DESIGN

> 日期：2026-07-06
> 实现：`Core/Security/SecretProtector.vb`、`Core/Security/ProtectedSettingsProvider.vb`。

## 1. 目标

对百度 **Secret Key** 做本地加密存储，避免以明文形式留存在 `user.config`。API Key、接口地址等非高度敏感项不加密。

## 2. 方案：Windows DPAPI（CurrentUser）

- 使用 `System.Security.Cryptography.ProtectedData`（.NET Framework 内置，无第三方依赖）。
- 作用域 `DataProtectionScope.CurrentUser`：密文仅当前 Windows 用户可解。
- 加密结果格式：`DPAPI:` + Base64(密文)。前缀用于识别与迁移。

## 3. 读写流程

| 场景 | 行为 |
|------|------|
| 保存 SK | `SecretProtector.Protect(明文)` → 存入 `My.Settings.BaiduSecretKey` |
| 读取 SK | `ProtectedSettingsProvider.GetSecretKey()` → `Unprotect` 解密 |
| 显示 SK（设置界面） | 加载时解密显示；输入框密码遮罩 |
| 使用 SK（OCR） | `OcrConfigProvider` 经 `ProtectedSettingsProvider` 解密后注入 `BaiduOcrOptions` |
| 解密失败 | 返回空串 + 标志；界面提示"无法解密，请重新输入"；**不抛出、不阻断插件加载** |
| 明文迁移 | 启动时 `MigratePlaintextIfNeeded()`：若存量为明文（无前缀）则加密回写 |

## 4. 安全边界（重要）

- **DPAPI 只保护"当前 Windows 用户 + 当前机器"下的本地数据**：换用户、换机器无法解密（此时提示重输）。
- **不能替代访问控制/权限管理**：能读取该用户 user.config 且以该用户身份运行的进程仍可解密。它降低"明文泄露"风险，而非提供强多方安全。
- 仅加密 Secret Key；API Key/URL/开关等普通配置不加密。
- **日志、文档、弹窗中不得出现明文 SK 或 token**；脱敏摘要仅保留前 2 后 2 字符。
- 兼容旧 XML（`baidu-ocr.config.xml`）中的明文仅作回退读取，建议改用设置界面并删除该文件（`.gitignore` 已忽略）。

## 5. 迁移方案

1. 旧版本把明文 SK 存于 `BaiduSecretKey`。
2. 新版本启动 `ThisAddIn_Startup` 调 `MigratePlaintextIfNeeded()`：检测无 `DPAPI:` 前缀 → 加密回写并 Save。
3. 之后所有读写均走加密路径。
4. 若迁移失败（异常），保持原值不变并记日志（不含明文），不影响加载。

## 5.1 配置文件落地（第七阶段）

- 除 My.Settings 外，新增**加密配置文件** `%AppData%\iWorkHelper\baidu-ocr.config.xml`（`Core/Configuration/BaiduXmlConfigStore.vb`）：
  - **Secret Key 以 DPAPI（CurrentUser）加密**（`DPAPI:` 前缀）；API Key/URL/开关明文。
  - `OcrConfigProvider` 在 My.Settings 未填 AK/SK 时**完整读取该文件并解密 SK**。
  - 该文件位于用户 `%AppData%`（**不在仓库**），`.gitignore` 已忽略 `baidu-ocr.config.xml`。
- 离线工具 `--save-baidu-config`（写入加密配置）/ `--use-config`（读取解密执行 OCR）。
- **实测（真实）**：写入后 `<SecretKey>DPAPI:...`、无 SK 明文；清空环境变量后 `--use-config` 解密并真实 OCR 成功 → 证明“加密落地→解密→使用”闭环。
- 安全提醒：AK 在该文件中为明文（Baidu 视 AK 为半公开、SK 为密钥）；文件受当前 Windows 用户 DPAPI 与文件系统权限保护，且不入库。

## 6. 已实现 / 已验证 / 未验证（第七阶段更新）

- **已实现**：加解密、前缀识别、失败兜底、明文迁移、界面/配置集成、加密配置文件落地。
- **已验证（真实执行）**：`OfflineTester --selftest` 在**当前 Windows 用户**下真实完成 DPAPI 往返：
  - 加密带 `DPAPI:` 前缀、密文≠明文、解密还原一致、明文遗留原样返回（迁移路径）。**5/5 通过。**
  - 此为与 VSTO 相同的 DPAPI 上下文（`DataProtectionScope.CurrentUser`）。
- **未验证**：Outlook 宿主进程内的保存→user.config 密文→重开解密→跨用户失败提示→明文自动迁移**整链路**（无 Outlook 环境）。见 [../testing/OUTLOOK_MANUAL_TEST.md](../testing/OUTLOOK_MANUAL_TEST.md)。

## 7. 后续可选增强

- 增加"清除已保存密钥"按钮。
- 对 API Key 也加密（当前视为较低敏感）。
- 支持企业级密钥管理（如从环境变量/密钥库读取），当前不引入以避免复杂化。
