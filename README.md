# Outlook MCP 服务器

一个基于模型上下文协议（MCP）的服务器，提供对Microsoft Outlook邮件功能的访问，允许大语言模型和其他MCP客户端通过标准化接口读取、搜索和管理邮件。

## 功能特性

- **文件夹管理**: 列出Outlook客户端中的所有可用邮件文件夹
- **邮件列表**: 获取指定时间段内的邮件
- **邮件搜索**: 通过联系人姓名、关键词或短语搜索邮件，支持OR操作符
- **邮件详情**: 查看完整的邮件内容，包括附件
- **撰写邮件**: 创建并发送新邮件
- **回复邮件**: 回复现有邮件

## 系统要求

- Windows操作系统
- Python 3.10或更高版本
- 已安装并配置Microsoft Outlook，且有活跃账户
- Claude Desktop或其他兼容MCP的客户端

## 安装步骤

1. 克隆或下载此仓库
2. 安装所需依赖：

```bash
pip install mcp>=1.2.0 pywin32>=305
```

3. 配置Claude Desktop（或您首选的MCP客户端）以使用此服务器

## 配置说明

### Claude Desktop配置

将以下内容添加到您的`MCP_client_config.json`文件中：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["您的路径\\outlook_mcp_server.py"],
      "env": {}
    }
  }
}
```

## 使用方法

### 启动服务器

您可以直接启动服务器：

```bash
python outlook_mcp_server.py
```

或者让Claude Desktop等MCP客户端通过配置自动启动。

### 可用工具

服务器提供以下工具：

1. `list_folders`: 列出Outlook中所有可用的邮件文件夹
2. `list_recent_emails`: 列出指定天数内的邮件标题
3. `search_emails`: 通过联系人姓名或关键词搜索邮件
4. `get_email_by_number`: 获取特定邮件的详细内容
5. `reply_to_email_by_number`: 回复特定邮件
6. `compose_email`: 创建并发送新邮件

### 使用流程示例

1. 使用`list_folders`查看所有可用的邮件文件夹
2. 使用`list_recent_emails`查看最近的邮件（例如最近7天）
3. 使用`search_emails`通过关键词查找特定邮件
4. 使用`get_email_by_number`查看完整邮件内容
5. 使用`reply_to_email_by_number`回复邮件

## 使用示例

### 列出最近邮件
```
请显示我最近3天的未读邮件
```

### 搜索邮件
```
搜索关于"项目更新 OR 会议纪要"的邮件，时间范围是最近一周
```

### 查看邮件详情
```
显示列表中第2封邮件的详细内容
```

### 回复邮件
```
回复第3封邮件："谢谢您的信息。我会审查这个内容，明天回复您。"
```

### 撰写新邮件
```
发送邮件给john.doe@example.com，主题是"会议议程"，内容是"这是我们即将召开会议的议程..."
```

## 故障排除

- **连接问题**: 确保Outlook正在运行且配置正确
- **权限错误**: 确保脚本有权限访问Outlook
- **搜索问题**: 对于复杂搜索，尝试在关键词之间使用OR操作符
- **邮件访问错误**: 检查邮件ID是否有效且可访问
- **服务器崩溃**: 检查Outlook的连接和稳定性

## 安全注意事项

此服务器可以访问您的Outlook邮件账户，能够读取、发送和管理邮件。请仅在受信任的MCP客户端和安全环境中使用。

## 功能限制

- 目前仅支持纯文本邮件（不支持HTML）
- 最大邮件历史记录限制为30天
- 搜索功能依赖于Outlook的内置搜索功能
- 仅支持基本邮件功能（不包括日历、联系人等）
