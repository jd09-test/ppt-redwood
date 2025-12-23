# Template-based PowerPoint Generator

基于Redwood模板的PowerPoint演示文稿自动生成器

现在市场上已经有很多AI生成PPT的项目（Gamma、Beautiful.ai等），这是一些非常优秀的项目，可以生成漂亮、丰富的演示文稿。
但是它们的创作的演示文稿样式基本由AI自动生成，无法严格遵循预定义格式，并且通常采用"先生成网页再转PPT"的方式。
虽然这些项目功能强大，但存在以下问题：
- **格式不统一**: 转换过程中容易丢失原始格式
- **样式不一致**: 难以保证企业品牌标准的严格遵循
- **模板限制**: 无法精确控制每个元素的位置和样式
- **不适宜编辑**: 生成的PPT文件对后续编辑不友好

本项目采用**严格遵循预定义模板**的方式创建PPT，确保：
- ✅ **品牌一致性**: 完全基于Oracle Redwood PPT模板，维护企业视觉标准
- ✅ **格式规范**: 通过预定义布局和占位符，保证格式统一
- ✅ **自动化生成**: 通过MCP协议，AI可以智能生成PPT的内容，而不必担心格式问题

## 功能特性

- 🎯 **模板驱动**: 严格基于Oracle Redwood PPT模板生成演示文稿
- 📐 **布局系统**: 支持多种预定义幻灯片布局模板
- 🎨 **富文本渲染**: 支持格式化文本内容渲染（粗体、斜体、下划线、颜色、超链接）
- 🌓 **主题切换**: 支持light/dark主题选择
- 🤖 **AI友好**: 通过MCP协议提供服务，便于AI调用


## 如何使用

### 前置条件

- Python and uv installed (docs: [MCP: Setting up your environment](https://modelcontextprotocol.io/quickstart/client#setting-up-your-environment))
- Coursor installed, 其它Agent客户端也许也可以，但是没有测试过 (docs: [Coursor: Installation](https://docs.cursor.com/en/get-started/installation))
- 下载本项目的代码，存放到合适的位置

### 在Coursor中配置MCP服务器

**注意：`WORK_DIR`是必须的，这是Cursor的工作目录，生成的文件都会保存在这个目录下**

- `directory`: 指定项目代码所在的目录
- `WORK_DIR`: 指定Cursor的工作目录

```json
{
  "mcpServers": {
    "ppt-generator": {
      "command": "uv",
      "args": ["--directory",
        "D:/MCP/templated-ppt-generator-mcp-server",
        "run",
        "src/server.py"
      ],
      "env": {"WORK_DIR": "D:/MCP/playground"}
    }
  }
}
```
### 使用MCP服务器生成PPT

直接在Coursor中输入你的需求，MCP服务器会根据你的需求生成PPT文件。

**Best Practices**
- 请确保你的需求是清晰、具体、完整的
- 提供恰当的输入信息，让AI了解背景知识
- AI总是避免不了幻觉，请一定要自己检查生成的PPT文件，确保没有错误

### 手动模式（无需Agent客户端）

自动化执行严重依赖AI Agent的能力，并非所有Agent客户端都能恰当地理解我们的需求。

因此我设计了手动模式（实际上我更喜欢手动执行多步操作，这样更可靠）。

- **动作 1**：生成提示文本，然后您可以使用任何您喜欢的AI来生成json内容。你需要把生成的内容保存为json文件。
- **动作 2**：从您生成的 JSON 文件生成演示文稿。提示您输入 `.json` 文件的路径，然后生成 `.pptx` 文件。

执行下面的脚本，执行手动模式。
```sh
uv run src/manual_run.py
```

### 局限性

- **模板依赖**: 需要Oracle PPT模板文件，并且要手工创建描述文件，模板文件必须存在于assets目录
- **布局限制**: 仅支持预定义的布局类型，无法动态创建新布局
- **占位符限制**: 仅支持预定义的占位符类型和位置
- **样式限制**: 仅支持基本的样式(加粗、斜体、下划线、颜色、超链接)，不支持复杂的CSS样式
- **多模态**：仅支持文本内容，不支持图片、图表等其他类型
- **AI Agent客户端**：仅在Cursor上做过测试验证，其他Agent客户端也许也可以（必须能够读写本地文件），但是没有测试过


## 技术原理

### 项目结构

```
ppt-generator/
└── src/                   # 源代码目录
    ├── model.py           # 数据模型定义
    ├── server.py          # MCP服务器实现
    ├── utils.py           # 功能函数
    └── assets/            # 资源文件
      └── Oracle_PPT-template_FY26_blank.pptx   # Oracle Redwood PPT模板
      └── layouts_template.yaml                 # 布局信息的描述文件
      └── prompt_template.txt                   # AI的提示词模板
```

### 技术栈

- **Python**: 核心开发语言
- **uv**: 依赖管理
- **FastMCP**: MCP服务器框架
- **python-pptx**: PowerPoint文件操作

### MCP服务器架构

项目基于FastMCP框架构建MCP服务器，提供以下工具：

- `get_presentation_rules`: 获取演示文稿生成规则
  - 包括对生成PPT的规则说明，如数据格式、样式、风格等
  - 包括预定义的PPT的布局描述，如标题、副标题、内容等，用json格式返回
- `generate_presentation`: 根据JSON内容,是要pptx python包生成PPT文件
  - 解析JSON内容，包括PPT的布局和内容
  - 根据JSON中的布局选择幻灯片布局
  - 解析JSON中的内容，包括样式和文本，填充占位符
  - 保存为PPTX文件

### AI工作流程

1. **理解需求**：AI 接收需求和背景信息，理解主题要求
2. **获取演示文稿规则**：调用 `get_presentation_rules` 工具获取生成规则，包括PPT的布局描述和输出格式等
3. **生成JSON内容**：AI根据规则和需求信息，创建JSON格式的演示文稿内容，其中包括布局、占位符和对应的样式、文本内容
4. **生成PPT文件**：调用 `generate_presentation` 工具，解析json内容，生成最终的PPT文件

### 模板系统

本项目使用Oracle Redwood PPT模板，模板文件必须存在于assets目录

- `Oracle_PPT-template_FY26_blank.pptx`：Oracle Redwood PPT模板，以此创建新文件，并且提取布局信息
- `layouts_template.yaml`：布局信息的描述文件，其中包含了手工创建的自然语言描述，用于生成AI的提示词
- `prompt_template.txt`：AI的提示词模板，其中包含了生成任务的要求，用于生成AI的提示词


