VERSION 5.00
Begin VB.UserForm frmAutoNest
   Caption         =   "自动化生产排版"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CheckBox chkRunNest
      Caption         =   "导入后执行批量生产+排版"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.TextBox txtCustomer
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
   End
   Begin VB.ComboBox cmbMaterial
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtCSV
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton cmdBrowse
      Caption         =   "..."
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   315
   End
   Begin VB.Label lblCustomer
      Caption         =   "客户名:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblMaterial
      Caption         =   "材料:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCSV
      Caption         =   "CSV 文件:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmAutoNest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    
    ' 材料列表（从 AD_MATERIALS 加载，方便扩展）
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
' 浏览按钮 — 选择 CSV 文件（预留，后期可接 CFileDialog）
' ==============================================================================
Private Sub cmdBrowse_Click()
    ' 简单输入框选择；后期可替换为 GetOpenFileName API
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
    
    ' 赋值输出参数
    CSVPath = Trim$(txtCSV)
    CustomerName = Trim$(txtCustomer)
    If CustomerName = "" Then CustomerName = DEF_CUSTOMER
    MaterialName = Trim$(cmbMaterial)
    RunNest = CBool(chkRunNest.Value)
    Cancelled = False
    
    ' 保存记忆
    SaveSetting "CCC", "AutoImportNest", "LastPath", CSVPath
    SaveSetting "CCC", "AutoImportNest", "Customer", CustomerName
    SaveSetting "CCC", "AutoImportNest", "Material", MaterialName
    
    Me.Hide
End Sub


' ==============================================================================
' 取消按钮
' ==============================================================================
Private Sub cmdCancel_Click()
    Cancelled = True
    Unload Me
End Sub
