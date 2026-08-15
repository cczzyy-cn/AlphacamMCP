# frmAutoNest 窗体 — 手动创建指南（CSV 文件选择版）

> AlphaCAM VBA **不支持导入 .frm 设计文件**，需手动创建窗体。
> 参照 `CCC功能/frmToolOffset.txt` 的方式保存代码。

## 创建步骤

1. 打开 AlphaCAM VBA 编辑器
2. 菜单：**插入 → 用户窗体(UserForm)** → 生成 `UserForm1`
3. 属性窗口：`(名称)` = **frmAutoNest**，`Caption` = **自动化生产排版**
4. 从工具箱添加 **8 个控件**（名称必须完全一致）：

| 控件 | 名称 | 标题 | 说明 |
|------|------|------|------|
| Label | `lblCSV` | "CSV 文件:" | 标签 |
| TextBox | `txtCSV` | — | CSV 路径 |
| CommandButton | `cmdBrowse` | "..." | 浏览（文件对话框） |
| CheckBox | `chkOnlyImport` | "只导入订单，不生产排版" | 勾选则跳过排版 |
| CheckBox | `chkOverwrite` | "强制覆盖重名订单（删除原订单所有数据）" | 勾选则重名时删除旧数据重导 |
| CommandButton | `cmdRegenLabel` | "重新生成标签" | 按当前嵌套图纸位置重新生成门板标签 EMF |
| CommandButton | `cmdOK` | "确定(&O)" | Default=True |
| CommandButton | `cmdCancel` | "取消(&C)" | Cancel=True |

5. 双击窗体空白处 → 代码窗口 → 全选删除默认代码 → 粘贴 [`frmAutoNest.txt`](frmAutoNest.txt) 全部代码
6. 编译保存（Ctrl+S）

## 布局示意

```
┌─ 自动化生产排版 ─────────────────────────┐
│  CSV 文件: [txtCSV      ] [...]          │
│  ☐ 只导入订单，不生产排版                │
│  ☐ 强制覆盖重名订单（删除原订单数据）    │
│  [cmdOK] [cmdCancel] [重新生成标签]      │
└────────────────────────────────────────┘
```

## 功能

- 点击"..." → 系统文件选择对话框选 CSV
- 勾选"只导入订单，不生产排版" → 仅导入订单（`bRunNest=False`），不调用 `g_Make_Master`
- 不勾选 → 导入 + 排版（`bRunNest=True`）
- 勾选"强制覆盖重名订单" → 订单名已存在时，删除原订单的 `AD_ORDER_DETAILS`、`AD_REPORT_DATA` 及订单本身后重新导入（`bOverwrite=True`）
- 不勾选且订单重名 → 提示并取消导入（原行为）
- 点击"重新生成标签" → ⚠️ **操作的是当前打开的图纸**，需先在 AlphaCAM 中打开排版档案 ARD 文件（如 `<订单>_<材料>.ard`，含嵌套信息）；否则会提示"当前图纸不是排版档案"并拒绝。手动移动/删除门板后使用：
  1. **备份当前图纸状态与原 ARD 文件**（临时副本，保护加工道次窗口不被破坏）
  2. 删除该订单/材料的旧件标签 EMF（清理残留）
  3. 按当前图纸位置重新生成 `<JobName>_<材料>_<板名>_<件号>.emf`
  4. **同步数据库**：按 `DEF_ATT_DETAIL_ID` 更新 `AD_REPORT_DATA.PressDoorImage/PressDoorCounter`，并删除已删板件的记录
  5. **恢复用户图纸**（重开备份副本）——原 ARD 文件内容不变，加工道次窗口状态不受影响
  （调用 `modAutoImportNest.g_RegenDoorLabelEMFs`）
- 窗体为**非模态**（不阻塞 AlphaCAM），操作完成后**保持打开**可连续导入多个订单；点"取消"才关闭
- 记忆上次路径（注册表）
