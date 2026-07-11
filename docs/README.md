# iWorkHelper 文档索引

**iWorkHelper** 是 Outlook VSTO 加载项（VB.NET / .NET Framework 4.8），用于批量识别邮件 PDF 附件（滴滴发票/行程单、常规增值税发票）并命名归档。本目录为项目文档集，以当前源码为准。

## 推荐阅读顺序

1. **项目总览** — [project/PROJECT_OVERVIEW.md](project/PROJECT_OVERVIEW.md)：定位、能力、状态、已知限制
2. **系统架构** — [architecture/CODE_STRUCTURE.md](architecture/CODE_STRUCTURE.md)、[architecture/TECH_STACK.md](architecture/TECH_STACK.md)
3. **核心处理流程** — [architecture/PROCESS_FLOW.md](architecture/PROCESS_FLOW.md)
4. **内外网构建策略** — [architecture/BUILD_VARIANTS.md](architecture/BUILD_VARIANTS.md)
5. **开发与识别细节** — development/ 下各文档
6. **测试** — testing/ 下各文档
7. **部署与故障排查** — deployment/ 下各文档
8. **历史与设计决策** — history/ 下各文档

## 文档清单

### project/ — 项目总览
| 文档 | 用途 |
|------|------|
| [PROJECT_OVERVIEW.md](project/PROJECT_OVERVIEW.md) | 项目定位、核心能力、验证状态、已知限制 |

### architecture/ — 架构与设计
| 文档 | 用途 |
|------|------|
| [CODE_STRUCTURE.md](architecture/CODE_STRUCTURE.md) | 代码目录结构、Core 模块职责、关键类 |
| [TECH_STACK.md](architecture/TECH_STACK.md) | 技术栈、依赖、构建/运行环境 |
| [PROCESS_FLOW.md](architecture/PROCESS_FLOW.md) | 批量归档端到端流程（分组→识别→分流→命名→归档）|
| [BUILD_VARIANTS.md](architecture/BUILD_VARIANTS.md) | 内网版/外网版条件编译、四套构建配置 |
| [BAIDU_OCR_INTEGRATION_DESIGN.md](architecture/BAIDU_OCR_INTEGRATION_DESIGN.md) | 百度 OCR 接入设计（token/请求/映射/调用策略）|
| [ERROR_HANDLING_DESIGN.md](architecture/ERROR_HANDLING_DESIGN.md) | 错误分类、预检查、运行锁、失败隔离、进度/报告 |

### requirements/ — 需求与规范
| 文档 | 用途 |
|------|------|
| [FIELD_EXTRACTION_SPEC.md](requirements/FIELD_EXTRACTION_SPEC.md) | 字段清单、本地/OCR 字段映射表 |
| [NAMING_TEMPLATE_SPEC.md](requirements/NAMING_TEMPLATE_SPEC.md) | 命名模板变量、取值优先级、fallback |
| [MAIL_PDF_CLASSIFICATION_SPEC.md](requirements/MAIL_PDF_CLASSIFICATION_SPEC.md) | 邮件/PDF 分流分类规则 |

### development/ — 识别与实现细节
| 文档 | 用途 |
|------|------|
| [GENERAL_INVOICE_RECOGNITION.md](development/GENERAL_INVOICE_RECOGNITION.md) | 常规发票本地识别（根因/方案/候选评分/明细/验证）|
| [LOCAL_TEXT_RECOGNITION_DIAGNOSTIC_REPORT.md](development/LOCAL_TEXT_RECOGNITION_DIAGNOSTIC_REPORT.md) | 滴滴行程单坐标行重建诊断与修复 |
| [OCR_ONLINE_VALIDATION_LOG.md](development/OCR_ONLINE_VALIDATION_LOG.md) | 百度 OCR 真实联调记录（含真实 bug 修复与字段校准）|
| [ARCHIVE_RUNNING_STATE_BUG_REPORT.md](development/ARCHIVE_RUNNING_STATE_BUG_REPORT.md) | "已有归档任务运行"自我阻断缺陷根因与修复 |
| [VERSION_DISPLAY_BUG_REPORT.md](development/VERSION_DISPLAY_BUG_REPORT.md) | 版本号显示修复 |

### security/ — 安全
| 文档 | 用途 |
|------|------|
| [SECRET_STORAGE_DESIGN.md](security/SECRET_STORAGE_DESIGN.md) | Secret Key DPAPI 加密方案与安全边界 |

### testing/ — 测试
| 文档 | 用途 |
|------|------|
| [TEST_PLAN.md](testing/TEST_PLAN.md) | 测试场景矩阵（离线 + 待 Outlook 人工）|
| [OFFLINE_TESTER_GUIDE.md](testing/OFFLINE_TESTER_GUIDE.md) | 离线工具用法、自测、样例脱敏与校准 |
| [REGRESSION_GUIDE.md](testing/REGRESSION_GUIDE.md) | 本地/常规发票回归方法 |
| [OUTLOOK_MANUAL_TEST.md](testing/OUTLOOK_MANUAL_TEST.md) | Outlook 端到端人工验证清单（待执行）|

### deployment/ — 部署与故障排查
| 文档 | 用途 |
|------|------|
| [USER_CONFIGURATION_GUIDE.md](deployment/USER_CONFIGURATION_GUIDE.md) | 用户配置与使用指南 |
| [TROUBLESHOOTING.md](deployment/TROUBLESHOOTING.md) | 常见问题排查 |
| [OUTLOOK_ADDIN_RESILIENCY_GUIDE.md](deployment/OUTLOOK_ADDIN_RESILIENCY_GUIDE.md) | Outlook 禁用加载项的诊断与恢复（用户侧）|
| [STARTUP_PERFORMANCE.md](deployment/STARTUP_PERFORMANCE.md) | 启动性能优化与慢启动根因诊断（开发/维护侧）|

### history/ — 历史与决策
| 文档 | 用途 |
|------|------|
| [DEVELOPMENT_HISTORY.md](history/DEVELOPMENT_HISTORY.md) | 开发历程（各阶段目标/关键功能/修复/遗留，整合自阶段汇报）|
| [DECISIONS_AND_RISKS.md](history/DECISIONS_AND_RISKS.md) | 关键设计决策、不可违背约束、仍有效风险 |

### archive/ — 历史归档（不代表当前实现）
接手初期的规划与快照文档，仅供追溯，见 [archive/README.md](archive/README.md)。
