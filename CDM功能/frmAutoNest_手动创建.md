# frmAutoNest 窗体 — 手动创建指南

> AlphaCAM VBA **不支持导入 .frm 设计文件**（VBE 报"不支持窗体类"）。
> 请按以下步骤手动创建窗体，然后粘贴代码。

## 创建步骤

1. 打开 AlphaCAM VBA 编辑器
2. 菜单：**插入 → 用户窗体(UserForm)** → 生成 `UserForm1`
3. 在属性窗口将 `(名称)` 改为 **frmAutoNest**，`Caption` 改为 **自动化生产排版**
4. 从工具箱添加以下控件（**名称必须完全一致**）：

| 控件 | 名称 | 类型 | 说明 |
|------|------|------|------|
| Label | lblCSV | 标签 | "CSV 文件:" |
| TextBox | txtCSV | 文本框 | CSV 路径 |
| CommandButton | cmdBrowse | 按钮 | "..." 浏览 |
| Label | lblMaterial | 标签 | "材料:" |
| ComboBox | cmbMaterial | 组合框 | 材料下拉 |
| Label | lblCustomer | 标签 | "客户名:" |
| TextBox | txtCustomer | 文本框 | 客户名 |
| CheckBox | chkRunNest | 复选框 | "导入后执行批量生产+排版"（勾选） |
| CommandButton | cmdOK | 按钮 | "确定(&O)" |
| CommandButton | cmdCancel | 按钮 | "取消(&C)" |

5. 双击窗体空白处打开代码窗口，**全选删除默认代码**，粘贴下方全部代码
6. 编译保存（Ctrl+S）

## 窗体代码（全部）

```vba
Option Explicit

' ==============================================================================
' 自动化生产排版 — 参数窗体
' ==============================================================================
' 提供参数输入，方便后期添加增强功能（材料/客户/选项等）
' 由 modAutoImportNest.AutoImportNest 调用
' ==============================================================================

' 输出参数（由调用模块读取）
Public CSVPath As String
Public CustomerName As String
Public MaterialName As String
Public RunNest As Boolean
Public Cancelled As Boolean

' 默认值
Private Const DEF_FOLDER As String = "C:\Users\C\Desktop\2026优化表\"
Private Const DEF_CUSTOMER As String = "自动化生产"


' ==============================================================================
' 窗体初始化 — 加载记忆值
' ==============================================================================
Private Sub UserForm_Initialize()
    Dim sLast As String
    Dim rst As Object
    Dim conn As Object
    
    Cancelled = True
    
    ' CSV 路径（记忆）
    sLast = GetSetting("CCC", "AutoImportNest", "LastPath", DEF_FOLDER & "7-10中林SPC婷兰灰.csv")
    txtCSV = sLast
    
    ' 客户名（记忆）
    txtCustomer = GetSetting("CCC", "AutoImportNest", "Customer", DEF_CUSTOMER)
    
    ' 材料列表（从 AD_MATERIALS 加载）
    cmbMaterial.Clear
    cmbMaterial.AddItem ""          ' 空 = 使用默认"开料机3000mm"
    On Error Resume Next
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\2016\LICOMDAT\CDM Data\CDM.mdb"
    Set rst = conn.Execute("SELECT Name FROM AD_MATERIALS ORDER BY Name")
    Do While Not rst.EOF
        cmbMaterial.AddItem rst.Fields("Name")
        rst.MoveNext
    Loop
    rst.Close
    conn.Close
    On Error GoTo 0
    
    ' 材料记忆
    cmbMaterial = GetSetting("CCC", "AutoImportNest", "Material", "")
End Sub


' ==============================================================================
' 浏览按钮 — 选择 CSV 文件
' ==============================================================================
Private Sub cmdBrowse_Click()
    Dim sInput As String
    sInput = InputBox("请输入 CSV 文件完整路径:", "选择 CSV", txtCSV)
    If sInput <> "" Then txtCSV = sInput
End Sub


' ==============================================================================
' 确定按钮
' ==============================================================================
Private Sub cmdOK_Click()
    If Trim$(txtCSV) = "" Then
        MsgBox "请选择 CSV 文件", vbExclamation, "自动化生产排版"
        Exit Sub
    End If
    If Dir(Trim$(txtCSV)) = "" Then
        MsgBox "文件不存在:" & vbCrLf & Trim$(txtCSV), vbExclamation, "自动化生产排版"
        Exit Sub
    End If
    
    ' 保存记忆
    SaveSetting "CCC", "AutoImportNest", "LastPath", Trim$(txtCSV)
    SaveSetting "CCC", "AutoImportNest", "Customer", Trim$(txtCustomer)
    SaveSetting "CCC", "AutoImportNest", "Material", Trim$(cmbMaterial)
    
    ' 调用模块带参入口（窗体不直接引用模块内部，由模块完成导入+排版）
    Me.Hide
    modAutoImportNest.AutoImportNestWithParams _
        Trim$(txtCSV), _
        Trim$(txtCustomer), _
        Trim$(cmbMaterial), _
        CBool(chkRunNest.Value)
    Unload Me
End Sub


' ==============================================================================
' 取消按钮
' ==============================================================================
Private Sub cmdCancel_Click()
    Cancelled = True
    Unload Me
End Sub
```

## 说明

- 创建后保存，`modAutoImportNest` 运行时会**自动检测**并使用该窗体
- 未创建窗体时自动回退 InputBox（现有功能不受影响）
- 后期增强：在窗体中加控件，在 `cmdOK_Click` 中赋值给输出属性即可
