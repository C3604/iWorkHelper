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

- **开发环境**：Visual Studio 2019 或以上，安装 VSTO 开发工具
- **依赖项**：
  - .NET Framework 4.7.2 或以上
  - Microsoft Office Outlook 2016 或以上
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

- **MailProcessor.vb**：邮件处理核心模块
- **PdfpigOcr.vb**：PDF 文本提取与 OCR 模块
- **MainRibbon.vb**：Outlook 功能区界面
- **SettingsForm.vb**：用户设置界面
- **LogManager.vb**：日志管理模块

## 配置说明

iWorkHelper 使用 `My.Settings` 存储配置信息，主要包括：

- **ArchivePath**：文件归档路径
- **MergeDidiFiles**：是否合并滴滴文件（发票+行程单）
- **其他自定义设置**

## 开发说明

### 扩展功能

如需扩展功能，可以关注以下几个方面：

1. **支持更多文档类型**：在 `ProcessPdfFiles()` 方法中添加新的文档类型处理逻辑
2. **增强 OCR 能力**：改进 `PdfpigOcr` 类中的文本识别算法
3. **自定义归档规则**：修改 `MovePdfFiles()` 方法中的归档逻辑

### 注意事项

- 避免在主线程中执行耗时操作
- 确保异常处理机制完善
- 优化大文件处理性能

## 贡献指南

欢迎提交 Pull Request 或 Issue，共同改进项目。贡献前请确保：

- 代码符合项目编码规范
- 提供完整的测试案例
- 更新相关文档

## 开源许可证

本项目采用 [MIT License](LICENSE.txt)。

## 联系方式

如有问题或建议，请通过 GitHub Issues 提交。

---

© 2025 iWorkHelper 