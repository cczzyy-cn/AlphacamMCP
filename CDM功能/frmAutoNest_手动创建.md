# frmAutoNest 窗体 — 手动创建指南（CSV 文件选择版）

> AlphaCAM VBA **不支持导入 .frm 设计文件**，需手动创建窗体。
> 参照 `CCC功能/frmToolOffset.txt` 的方式保存代码。

## 创建步骤

1. 打开 AlphaCAM VBA 编辑器
2. 菜单：**插入 → 用户窗体(UserForm)** → 生成 `UserForm1`
3. 属性窗口：`(名称)` = **frmAutoNest**，`Caption` = **自动化生产排版**
4. 从工具箱添加 **6 个控件**（名称必须完全一致）：

| 控件 | 名称 | 标题 | 说明 |
|------|------|------|------|
| Label | `lblCSV` | "CSV 文件:" | 标签 |
| TextBox | `txtCSV` | — | CSV 路径 |
| CommandButton | `cmdBrowse` | "..." | 浏览（文件对话框） |
| CheckBox | `chkOnlyImport` | "只导入订单，不生产排版" | 勾选则跳过排版 |
| CommandButton | `cmdOK` | "确定(&O)" | Default=True |
| CommandButton | `cmdCancel` | "取消(&C)" | Cancel=True |

5. 双击窗体空白处 → 代码窗口 → 全选删除默认代码 → 粘贴 [`frmAutoNest.txt`](frmAutoNest.txt) 全部代码
6. 编译保存（Ctrl+S）

## 布局示意

```
┌─ 自动化生产排版 ────────────────┐
│  CSV 文件: [txtCSV      ] [...] │
│  ☐ 只导入订单，不生产排版       │
│            [cmdOK] [cmdCancel]  │
└───────────────────────────────┘
```

## 功能

- 点击"..." → 系统文件选择对话框选 CSV
- 勾选"只导入订单，不生产排版" → 仅导入订单（`bRunNest=False`），不调用 `g_Make_Master`
- 不勾选 → 导入 + 排版（`bRunNest=True`）
- 记忆上次路径（注册表）
