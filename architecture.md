# 代码架构文档

## 1. 核心文件结构

| 文件名称 | 功能描述 | 主要职责 |
|---------|---------|----------|
| `src/weekly_report_sender.py` | 核心功能实现 | 协调各模块工作，执行主工作流程 |
| `src/config/config_manager.py` | 配置管理 | 加载和管理配置 |
| `src/excel/excel_reader.py` | Excel处理 | 读取Excel文件和转换格式 |
| `src/html/html_generator.py` | HTML生成 | 生成HTML内容 |
| `src/email/email_sender.py` | 邮件发送 | 发送邮件 |
| `config.py` | 配置文件 | 存储邮件服务器、收件人、文件路径等配置 |
| `requirements.txt` | 依赖管理 | 列出项目依赖项 |
| `README.md` | 项目文档 | 项目说明和使用指南 |
| `weekly_report.xlsx` | 示例Excel文件 | 周报数据示例 |

## 2. 模块架构

| 模块 | 文件 | 功能描述 | 职责 |
|-----|------|---------|------|
| 配置管理 | `src/config/config_manager.py` | 加载和管理配置 | 从环境变量和配置文件加载配置 |
| Excel处理 | `src/excel/excel_reader.py` | 读取Excel文件和转换格式 | 读取Excel内容，处理合并单元格，转换为HTML |
| HTML生成 | `src/html/html_generator.py` | 生成HTML内容 | 生成带样式的HTML邮件内容 |
| 邮件发送 | `src/email/email_sender.py` | 发送邮件 | 创建和发送邮件 |
| 主协调器 | `src/weekly_report_sender.py` | 协调各模块工作 | 执行主工作流程 |

## 3. 主要类和方法

| 类/方法 | 功能描述 | 代码行数 | 复杂度 |
|---------|---------|----------|--------|
| WeeklyReportSender | 周报发送主类 | 492 | 高 |
| validate_config | 验证配置 | 26 | 低 |
| read_excel_with_merged_cells | 处理合并单元格 | 65 | 中 |
| parse_excel_structure | 解析Excel结构 | 30 | 中 |
| read_excel_content | 读取Excel内容并生成HTML | 123 | 高 |
| create_email_message | 创建邮件消息 | 26 | 低 |
| send_email | 发送邮件 | 35 | 中 |
| run | 主执行方法 | 21 | 低 |

## 4. 系统流程

1. **配置加载**：从环境变量和配置文件加载配置信息
2. **Excel读取**：读取Excel文件内容，处理合并单元格，保留格式信息
3. **HTML生成**：将Excel内容转换为HTML格式，保留样式
4. **邮件创建**：创建包含HTML内容的邮件消息
5. **邮件发送**：通过SMTP发送邮件到配置的收件人和抄送人
6. **日志记录**：记录系统运行状态和错误信息

## 5. 技术栈

- **Python 3**：主要开发语言
- **openpyxl**：读取Excel文件和处理格式
- **smtplib**：发送邮件
- **email**：创建邮件消息
- **python-dotenv**：支持环境变量配置
- **logging**：系统日志记录

## 6. 配置管理

系统支持两种配置方式：

1. **环境变量**：通过`.env`文件或系统环境变量配置
2. **配置文件**：通过`config.py`文件配置

配置项包括：
- 邮件服务器设置（SMTP服务器、端口、用户名、密码）
- 收件人设置（收件人、抄送人）
- 文件设置（Excel文件路径、邮件主题）

## 7. 安全性

- 敏感信息（如邮件密码）从环境变量读取，避免硬编码
- 支持TLS/SSL加密的SMTP连接
- 输入验证和错误处理

## 8. 可扩展性

- 模块化设计，便于添加新功能
- 清晰的接口定义
- 易于测试和维护

## 9. 性能优化

- 移除了不必要的依赖
- 优化了Excel处理逻辑
- 减少了内存使用

## 10. 测试

- 测试脚本：
  - `test/test_system.py` - 测试完整系统功能
  - `test/test_config.py` - 测试配置管理
  - `test/test_excel.py` - 测试Excel读取，包括随机增加行和列
  - `test/test_html.py` - 测试HTML生成
  - `test/test_email.py` - 测试邮件发送
  - `test/test_main.py` - 测试主协调器
- 测试覆盖：
  - 配置加载和验证
  - Excel文件读取和处理
  - 合并单元格处理
  - 格式保留
  - HTML生成
  - 邮件创建和发送
  - 随机增加行和列的处理