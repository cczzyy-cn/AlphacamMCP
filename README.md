# AlphaCAM MCP 桥接器 — AI 驱动的 CAM 自动化

通过 [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) 让 AI 助手直接操作 **AlphaCAM 2016 R1**。

## 概述

本桥接器将 AlphaCAM 2016 R1 的 COM API 封装为 MCP 工具，使 AI 能够实时：
- 创建/读取/修改几何图形（线、圆、矩形、多边形、椭圆）
- 执行加工操作（粗精加工、轮廓铣槽、雕刻、钻孔等）
- 管理刀具和刀具路径
- 处理排版嵌套（Nesting）
- 运行 VBA 宏
- 输出 NC 代码

## 文件说明

| 文件 | 说明 |
|---|---|
| `server.py` | MCP 桥接器主程序（Python），通过 STDIO 协议与 AI 通信 |
| `alphacam_com.py` | AlphaCAM COM 自动化封装层 |
| `CCC功能合集.bas` | VBA 插件模块（依边界裁剪、全排版刀具偏移、排版刀具排序） |
| `frmToolOffset.txt` | 刀具偏移对话框的窗体定义（VBA UserForm） |
| `frmToolSort.txt` | 刀具排序对话框的窗体定义（VBA UserForm） |
| `install.bat` | Windows 一键安装脚本 |
| `install_vba.py` | VBA 代码安装到 AlphaCAM 的 Python 脚本 |
| `make_icons.py` | 生成工具栏 BMP 图标的工具 |
| `SKILL.md` | MCP 技能定义 |
| `requirements.txt` | Python 依赖 |

## 安装

### 前提条件

- Windows 7+ / 10 / 11
- AlphaCAM 2016 R1 已安装
- Python 3.10+（需要 `pywin32` 和 `mcp` 库）

### 快速安装

```bash
# 安装 Python 依赖
pip install -r requirements.txt

# 双击运行 install.bat，或手动注册到 AI 客户端的 MCP 配置
```

### 手动配置

在 AI 客户端的 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "alphacam-bridge": {
      "command": "python",
      "args": [
        "C:\path\to\server.py",
        "--progid",
        "aroutaps.Application"
      ]
    }
  }
}
```

## 可用功能

### 状态与信息
- 检查 AlphaCAM 连接状态
- 获取当前图纸信息（几何、刀具路径、操作、图层）

### 几何创建
- 矩形、圆、直线、多边形、椭圆、文本

### 加工操作
- 粗精加工、轮廓铣槽、雕刻、钻孔、啄钻、攻丝、镗孔
- 完整的进给/转速/深度控制

### 刀具与路径
- 选择刀具、获取当前刀具信息
- 列出/管理操作和刀具路径

### 排版嵌套
- 读取 NestInformation
- 遍历 Sheet 和 Part

### VBA 与插件
- 运行 VBA 宏
- 加载/启用插件

### 实用工具
- 设置撤销点、缩放全图、选择后处理器
- 批量工作流（`run_workflow`）

## VBA 插件功能

`CCC功能合集.bas` 包含三个工具（通过 AlphaCAM 菜单栏 "CCC功能" 访问）：

1. **依边界裁剪** — 选择边界和线段，将线段超出边界的部分裁剪
2. **全排版刀具偏移** — 按刀具名称选择，整体偏移 X/Y/Z
3. **排版刀具排序** — 按加工方式+刀具分组，拖拽调整加工顺序

## 项目链接

- GitHub: https://github.com/cczzyy-cn/AlphacamMCP
- 问题反馈: https://github.com/cczzyy-cn/AlphacamMCP/issues
