# Marp Fast PPT Maker - MCP Server

一个基于 MCP（Model Context Protocol）的 AI 驱动幻灯片生成服务器，快速将 Markdown 内容转换为专业 PPTX 演示文稿。

## 功能特性

- ✅ **快速生成**：从 Markdown 自动生成 PPTX 幻灯片
- ✅ **多主题支持**：default、gaia、uncover 三种专业主题
- ✅ **智能分页**：自动识别标题进行分页
- ✅ **本地文件支持**：支持插入本地图片
- ✅ **完整日志**：详细的调试信息便于问题排查
- ✅ **超时保护**：防止无限等待的 30 秒超时机制

## 快速开始

### 1. 安装依赖

**Python 依赖：**
```bash
python -m venv venv
source venv/bin/activate  # macOS/Linux
pip install -r requirements.txt
```

**Node.js 依赖：**
```bash
npm install
```

### 2. 测试服务器

运行直接测试检验 marp 命令：
```bash
python test_marp_direct.py
```

### 3. 配置 MCP 客户端

#### 使用 Claude Desktop

复制配置文件到 Claude Desktop 配置目录：

**macOS/Linux：**
```bash
cp claude_desktop_config.json \
   ~/Library/Application\ Support/Claude/claude_desktop_config.json
```

**Windows：**
```powershell
Copy-Item claude_desktop_config.json -Destination "$env:APPDATA\Claude\claude_desktop_config.json"
```

然后重启 Claude Desktop。

#### 使用 MCP Inspector（开发调试）

```bash
npx @modelcontextprotocol/inspector \
  python -u server.py
```

然后打开 http://localhost:5173

## 工具使用

### create_presentation 工具

**参数：**

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `title` | string | 必需 | 输出文件名（不带后缀） |
| `content` | string | 必需 | Markdown 幻灯片内容 |
| `theme` | enum | `default` | 选择主题：`default`、`gaia`、`uncover` |
| `style_class` | string | `lead` | 样式类：`lead` (居中) 或 `invert` (反色) |
| `auto_split` | boolean | `true` | 自动在标题处分页 |

**示例：**

```python
create_presentation(
    title="quarterly_report",
    content="""# Q1 Report

## Sales
- 300K this quarter
- 20% growth

## Team
- Expanded to 50 people
""",
    theme="gaia",
    auto_split=True
)
```

## 输出

生成的 PPTX 文件存储在 `output_slides/` 目录中。

## 调试

服务器会输出详细的调试信息到 `stderr`，包括：

- `[STARTUP]` - 服务器启动信息
- `[DEBUG]` - 处理过程的详细日志
- `[ERROR]` - 错误信息和堆栈跟踪

要查看实时日志（使用 Claude Desktop）：
```bash
tail -f ~/Library/Logs/Claude/mcp*.log
```

## 常见问题

**Q: 工具未在 Claude Desktop 中显示？**
- 检查配置文件路径是否正确
- 确认虚拟环境路径在配置中是绝对路径
- 重启 Claude Desktop

**Q: 生成 PPTX 失败或超时？**
- 检查 Marp CLI 是否安装：`npm list -g @marp-team/marp-cli`
- 查看服务器日志中的 `[ERROR]` 信息
- 确保有足够的磁盘空间

**Q: 怎样修改主题？**
在 `create_presentation` 的 `theme` 参数中选择：
- `default` - 白底专业风格
- `gaia` - 彩色深色，适合演讲
- `uncover` - 极简大字体

## 项目结构

```
marp-fast_ppt_maker-mcp/
├── server.py                      # MCP 服务器主程序
├── requirements.txt               # Python 依赖
├── package.json                   # Node.js 依赖
├── claude_desktop_config.json     # Claude Desktop 配置模板
├── .gitignore                     # Git 忽略文件
├── README.md                      # 本文件
├── SETUP_GUIDE.md                 # 详细设置指南
├── test_marp_direct.py           # 直接测试 marp 命令
└── output_slides/                 # 输出目录（运行时生成）
```

## 技术栈

- **框架**：MCP (Model Context Protocol)
- **Python 版本**：3.11+
- **幻灯片引擎**：Marp CLI
- **通信协议**：STDIO JSON-RPC

## 相关链接

- [MCP 官方文档](https://modelcontextprotocol.io/)
- [Marp 官方网站](https://marp.app/)
- [Claude 官方文档](https://docs.anthropic.com/)

## 许可证

MIT

## 支持

如有问题或建议，欢迎提交 Issue 或 Pull Request。
