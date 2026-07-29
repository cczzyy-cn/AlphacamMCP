# AI工作流程 — CDM 自动化工件

## 文件说明

| 文件 | 来源 | 说明 |
|------|------|------|
| `BatchImport.bas` | 项目 `CDM功能/` | CSV 导入→创建订单（自定义模块，已实现） |
| `BatchProcess.bas` | 项目 `CDM功能/` | 批量生产生成几何（自定义模块，已实现） |
| `CDM右键菜单事件位置.txt` | CDM.arb 源码分析 | frmNTCW 中右键菜单对应的事件处理位置说明 |

## 完整工作流

```python
# ═══════════════════════════════════════════════════════
# 三步完全自动化
# ═══════════════════════════════════════════════════════

# 第 1 步：安装模块（仅首次）
code = open("CDM功能/BatchImport.bas").read()
install_vba_module(module_name="BatchImport", code=code)
code = open("CDM功能/BatchProcess.bas").read()
install_vba_module(module_name="BatchProcess", code=code)

# 第 2 步：导入 CSV → 创建订单
run_vba_macro(
    macro_name="CDM.BatchImport.Run",
    params=["C:\\Users\\C\\Desktop\\2026优化表\\7-10中林SPC婷兰灰.csv",
            "7-10中林SPC婷兰灰"],
)

# 第 3 步：批量生产
run_vba_macro(
    macro_name="CDM.BatchProcess.RunByName",
    params=["7-10中林SPC婷兰灰"],
)

# 排序 + NC 输出
run_vba_macro(macro_name="CDM.m_OrderToolPaths")
select_post(post_name="T3香蕉弯")
output_nc(file_path="D:\\output\\7-10中林SPC婷兰灰.nc")
```

## 原始 CDM 右键菜单代码位置

CDM 主界面 frmNTCW（132KB VBA 窗体代码）中：
- 右键菜单事件在 **窗体中段**（代码截断部分，约 60KB-99KB 区域）
- 处理函数名未知（无法完整读取）
- 替代方案：使用上述自定义模块实现相同功能

---

## BatchImport 导入字段映射

与 CDM 原版 CSV 导入配置一致（1-based 列号）：

| CDM 字段 | 1-based 列号 | 0-based 索引 | CSV 实际列名 |
|---------|:-----------:|:-----------:|-------------|
| 门板类型 | 列1 | 0 | 造型名称 |
| 宽 | 列2 | 1 | 宽度 |
| 高 | 列3 | 2 | 高度 |
| 数量 | 列4 | 3 | 数量 |
| 组编号 | 列5 | 4 | 颜色 |
| 客户名 | 列6 | 5 | 客户名称 |
| 订单号 | 列10 | 9 | 板件码 |
| 生产注释 | 列12 | 11 | 备注 |
| **材料** | **列13 (M列)** | **12** | **(空 → 默认材料)** |

### 材料处理

- **来源列**：第13列（M列），0-based 索引 `COL_MATERIAL = 12`
- **默认材料**：当 CSV 中第13列为空时，使用 `"开料机3000mm"`（18×1220×3000mm）
- **自动创建**：材料名在 AD_MATERIALS 中不存在时自动创建

### 常量定义（BatchImport.bas）

```vb
Private Const COL_STYLE_NAME    As Integer = 0   ' 列1: 门板类型
Private Const COL_WIDTH         As Integer = 1   ' 列2: 宽
Private Const COL_HEIGHT        As Integer = 2   ' 列3: 高
Private Const COL_QUANTITY      As Integer = 3   ' 列4: 数量
Private Const COL_GROUP_ID      As Integer = 4   ' 列5: 组编号
Private Const COL_CUSTOMER      As Integer = 5   ' 列6: 客户名
Private Const COL_ORDER_REF     As Integer = 9   ' 列10: 订单号
Private Const COL_REMARK        As Integer = 11  ' 列12: 生产注释
Private Const COL_MATERIAL      As Integer = 12  ' 列13: 材料

Private Const DEF_MATERIAL_NAME As String = "开料机3000mm"
Private Const DEF_MATERIAL_THK  As Double = 18
Private Const DEF_MATERIAL_W    As Double = 1220
Private Const DEF_MATERIAL_L    As Double = 3000
```
