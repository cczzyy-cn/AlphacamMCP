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
