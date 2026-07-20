' ==============================================================================
' CCC功能 - Events 入口模块
' ==============================================================================
Option Explicit
Option Private Module

Function InitAlphacamAddIn(AcamVersion As Long) As Integer
    Dim frm As Frame: Set frm = App.Frame
    With frm
        Dim barId As Long: barId = .CreateButtonBar("CCC功能")
        .AddMenuItem3 "依边界裁&剪", "m_依边界裁剪", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "cut.bmp", .LastMenuCommandID
        .AddMenuItem3 "全排版刀具偏&移", "m_全排版刀具偏移", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "offset.bmp", .LastMenuCommandID
        .AddMenuItem3 "排版刀具排&序", "m_排版刀具排序", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "sort.bmp", .LastMenuCommandID
        .AddMenuItem3 "反面镜&像", "m_反面镜像", acamMenuNEW, "CCC功能", vbNullString
        .AddMenuItem3 "斜角下&刀", "m_斜角下刀", acamMenuNEW, "CCC功能", vbNullString
        .AddMenuItem3 "路径自身镜&像", "m_路径自身镜像", acamMenuNEW, "CCC功能", vbNullString
        .AddMenuItem3 "测试路径信息", "m_测试选中路径信息", acamMenuNEW, "CCC功能", vbNullString
    End With
    InitAlphacamAddIn = 0
End Function

Function m_依边界裁剪(): 依边界裁剪: End Function
Function m_全排版刀具偏移(): 全排版刀具偏移: End Function
Function m_排版刀具排序(): 排版刀具排序: End Function
Function m_反面镜像(): 反面镜像: End Function
Function m_斜角下刀(): 斜角下刀: End Function
Function m_路径自身镜像(): 路径自身镜像: End Function
Function m_测试选中路径信息(): 测试选中路径信息: End Function
