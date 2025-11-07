# iWorkHelper - Outlook邮件处理助手

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 项目简介

iWorkHelper 是一个基于 VSTO（Visual Studio Tools for Office）和 VB.NET 的 Outlook 插件，旨在自动化处理邮件附件中的 PDF 文件，特别是针对滴滴出行相关的发票和行程单。该插件集成了 OCR 文本识别功能，可以智能提取文档信息，并实现自动化的文件重命名和归档管理。

## 功能特点

- **邮件附件处理**：自动提取并保存邮件中的附件
- **OCR文本识别**：利用 PdfPig 库从 PDF 文档中提取文本信息
- **滴滴文档合并**：自动识别并合并滴滴出行的发票和行程单
- **智能文件重命名**：基于文档内容自动生成规范化的文件名
- **自动文件归档**：将处理后的文件移动到指定归档目录
- **邮件状态标记**：处理完成后自动标记邮件状态
- **进度可视化**：提供处理进度的实时显示界面
- **自定义配置**：支持用户自定义归档路径和处理参数

## 环境要求

- **开发环境**：Visual Studio Community 2022 (64 位)  
  - 安装 VSTO 开发工具
- **依赖项**：
  - .NET Framework 8 或以上
  - Outlook：Office 365 版本 2502（内部版本 18526.20604） 
  - PdfPig 库（PDF 文本提取和 OCR）
  - iTextSharp 库（PDF 文件合并）

## 安装指南

1. **克隆仓库**
   ```
   git clone https://github.com/C3604/iWorkHelper.git
   ```

2. **使用 Visual Studio 打开解决方案**
   ```
   XYOutlookPlugin.sln
   ```

3. **安装依赖项**
   - 通过 NuGet 包管理器安装所需依赖
   - 确保 VSTO 开发工具已正确安装

4. **编译和部署**
   - 在 Visual Studio 中编译解决方案
   - 使用 ClickOnce 发布方式进行部署

## 使用说明

1. 在 Outlook 中打开要处理的邮件
2. 在功能区找到 iWorkHelper 插件按钮
3. 点击"处理邮件"按钮开始自动化处理
4. 系统将自动执行以下操作：
   - 提取邮件附件
   - 执行 OCR 识别
   - 重命名文件（基于内容识别）
   - 归档到指定目录
   - 标记邮件状态

## 项目结构


## 配置说明

iWorkHelper 使用 `My.Settings` 存储配置信息，主要包括：

### SettingsWriter 使用

- 新增/更新配置项（仅支持 `Boolean` 与 `String`）：

  ```vb
  ' 将字符串类型配置写入（存在则更新，不存在则创建）
  Dim r1 = SettingsWriter.WriteSetting("ArchivePath", "String", "C:\\Archive")
  ' 将布尔类型配置写入
  Dim r2 = SettingsWriter.WriteSetting("EnableFeatureX", "Boolean", True)

  ' 返回值：Tuple(Of Boolean, String)
  ' r1.Item1=True 表示成功；失败时 r1.Item2 为错误原因
  ```

- 错误示例：
  - 类型不匹配：`SettingsWriter.WriteSetting("X", "Boolean", "True")`
  - 类型无效：`SettingsWriter.WriteSetting("X", "Integer", 1)`
  - 变量名无效（空/不合法标识符）：`SettingsWriter.WriteSetting("", "String", "v")`

### 日志系统使用（LogManager）

- 日志路径：`{tmppath}\log\iWorkhelper.log`，其中 `{tmppath}` 来源于 `My.Settings.tmppath`（未配置时回退到系统临时目录）。
- 统一 UTF-8 编码保存；单个日志超过 10MB 自动分割。
- 日志级别：`Info/Warn/Error`；`Warn` 与 `Error` 会同步弹窗（分别为黄色警告与红色错误）。

示例：

```vb
Imports LogManager

' 在加载时初始化
Dim logger As New Logger() ' 使用 My.Settings.tmppath
logger.Start()

' 记录常规信息
logger.LogInfo("启动成功")

' 记录警告（同步弹窗）
logger.LogWarn("网络波动，请稍后重试")

' 记录错误（同步弹窗）
logger.LogError("数据库连接失败")

' 退出时停止
logger.Stop()
```

注意：
- 为保证一致性，其他模块不应直接调用 `MessageBox.Show`；应统一通过 `Logger.LogWarn/LogError` 触发弹窗与日志。
- 写入操作在后台线程批量进行，避免阻塞主线程；必要时可调整构造参数中的刷新间隔与批量大小。


## 开发说明

### 扩展功能

如需扩展功能，可以关注以下几个方面：

### 注意事项

- 避免在主线程中执行耗时操作
- 确保异常处理机制完善
- 优化大文件处理性能


## 开源许可证

本项目采用 [MIT License](LICENSE.txt)。



---

© 2025 iWorkHelper