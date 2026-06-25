' ==============================================================================
' CCCFUNC — modSort 排版刀具排序
' ==============================================================================
' 功能：对排版图纸中所有加工操作，按用户指定的刀具+加工方式顺序重新排列。
'       支持多 Sheet 排版，按 Sheet 分组排序，每个 Sheet 内的操作按用户顺序排列。
'
' 核心流程：
'   排版刀具排序() → 弹出 frmToolSort 对话框
'      ├─ ScanOperations()     → 扫描图纸，收集所有不重复的"加工方式+刀具"组合
'      ├─ 用户拖拽调整顺序
'      └─ ApplySortToDrawing() → 按用户顺序重排各 Sheet 的刀具路径
'           ├─ 建立 g_mapPathToSheet：刀具路径名 → Sheet 编号
'           ├─ 建立 stD 分组字典：Sheet|加工方式+刀具 → 刀具路径集合
'           ├─ 清空所有 OpNo
'           └─ 按 Sheet + 用户排序重设 OpNo → OrderAll 生效
'
' 关键数据结构：
'   g_mapPathToSheet  — 全局字典，缓存 路径名 → Sheet编号 的映射
'   shNm()            — 数组(每个Sheet一个字典)，记录该Sheet上的所有路径名
'   stD               — 字典，"SheetID|加工方式 T刀号 刀具名" → Collection(Path)
'   sortedKeys()      — 用户从对话框拖拽确定的操作顺序列表
'   Path.OpNo         — AlphaCAM 路径排序号，格式 Sheet*100+pos
' ==============================================================================
Option Explicit
Option Private Module

' --------------------------------------------------------------------------
' 全局字典：路径名 → Sheet编号
' 缓存避免重复扫描 NestInformation，仅首次调用时初始化
' --------------------------------------------------------------------------
Private g_mapPathToSheet As Object  ' Scripting.Dictionary
Private g_lastDrawingSig As String  ' 上次排序的图纸签名，检测图纸变更用


' ==============================================================================
' Sub 排版刀具排序()
' ==============================================================================
' 入口过程。由 AlphaCAM 菜单"CCC功能 > 排版刀具排序"触发。
' 检查图纸和操作是否存在，然后弹出排序对话框。
' ==============================================================================
Sub 排版刀具排序()
    ' 每次打开对话框时清空缓存，强制重新扫描（图纸可能已变更）
    Set g_mapPathToSheet = Nothing
    
    Dim drw As Drawing
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "排版刀具排序"
        Exit Sub
    End If
    
    If drw.Operations Is Nothing Or drw.Operations.Count = 0 Then
        MsgBox "图纸中没有加工操作！", vbExclamation, "排版刀具排序"
        Exit Sub
    End If
    
    ' vbModeless — 非模态显示，用户可在对话框打开时操作 AlphaCAM
    frmToolSort.Show vbModeless
End Sub


' ==============================================================================
' Function ScanOperations() As Collection
' ==============================================================================
' 扫描图纸中所有 Operation → SubOperation，收集不重复的加工方式+刀具组合。
'
' 返回值：Collection，每项是字符串 "加工方式 T刀号 刀具名"
'         例如 "粗加工 T1 6mm Endmill", "钻孔 T3 3mm Drill"
'
' 去重逻辑：同一个加工方式+同一把刀只出现一次，因为用户排序是以
'           "加工方式-刀具"为单位进行整组移动，而非单个路径。
'
' 注意：所有 Dim 声明已移至函数顶部，避免 AlphaCAM VBA 编译报错。
' ==============================================================================
Public Function ScanOperations() As Collection
    Dim result As New Collection
    Dim drw As Drawing
    Dim ops As Operations
    Dim dict As Object
    Dim i As Long, j As Long
    Dim op As Operation
    Dim subs As SubOperations
    Dim subop As SubOperation
    Dim t As MillTool
    Dim methodName As String
    Dim spPos As Integer
    Dim key As String
    Dim keysArr As Variant
    Dim kk As Long
    
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        Set ScanOperations = result
        Exit Function
    End If
    
    Set ops = drw.Operations
    If ops Is Nothing Then
        Set ScanOperations = result
        Exit Function
    End If
    
    ' 用 Dictionary 做去重（VBA 没有 Set/哈希集合）
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To ops.Count                                  ' 遍历所有 Operation
        Set op = ops(i)
        
        Set subs = op.SubOperations
        If subs Is Nothing Then GoTo NextOp2                  ' 无子操作则跳过
        
        For j = 1 To subs.Count                              ' 遍历每个 SubOperation
            Set subop = subs(j)
            
            Set t = subop.Tool
            If Not (t Is Nothing) Then
                ' 提取加工方式名称（去掉后面的空格和细节描述）
                ' 例如 "粗加工  零件" → "粗加工"
                methodName = subop.Name
                
                ' 先尝试按双空格截断，再试单空格
                spPos = InStr(methodName, "  ")               ' 查找双空格
                If spPos > 0 Then
                    methodName = Left(methodName, spPos - 1)
                Else
                    spPos = InStr(methodName, " ")
                    If spPos > 0 Then methodName = Left(methodName, spPos - 1)
                End If
                
                ' 构造唯一键：加工方式 + 刀号 + 刀具名
                ' 格式如 "粗加工 T1 6mm Endmill"
                key = methodName & " T" & CStr(t.Number) & " " & t.Name
                
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        Next j
NextOp2:
    Next i
    
    ' 将 Dictionary 的 Keys 转为 Collection 返回
    keysArr = dict.Keys
    For kk = 0 To UBound(keysArr)
        result.Add keysArr(kk)
    Next kk
    
    Set ScanOperations = result
End Function


' ==============================================================================
' Sub ApplySortToDrawing(ByRef sortedKeys() As String)
' ==============================================================================
' 应用用户指定的排序顺序到图纸。
'
' sortedKeys() — 用户从对话框拖拽确定的顺序数组，每项为 "加工方式 T刀号 刀具名"
'
' 算法分三个阶段：
'   阶段1 — 建立 g_mapPathToSheet（路径名→Sheet映射，只做一次）
'   阶段2 — 遍历所有 SubOperation，按 "SheetID|加工方式+刀具" 分组收集刀具路径
'   阶段3 — 清空旧 OpNo → 按 Sheet×顺序 重设 OpNo → OrderAll 排序
'
' OpNo 编码规则：Sheet编号×100 + 组内序号
'   例如 Sheet 1 的第 3 组 → OpNo = 103
'   OrderAll 按 OpNo 升序重排全部刀具路径
'
' 注意：所有 Dim 声明已移至函数顶部，避免 AlphaCAM VBA 编译报错。
' ==============================================================================
Public Sub ApplySortToDrawing(ByRef sortedKeys() As String)
    On Error GoTo ErrHandler3
    
    ' ==================================================================
    ' 集中声明区 — 所有变量在此声明，不再分散在 For/If 块内
    ' ==================================================================
    
    ' --- 主对象 ---
    Dim drw As Drawing
    Dim ops As Operations
    Dim ni As NestInformation
    Dim ni2 As NestInformation
    
    ' --- 循环索引 ---
    Dim opIdx As Long, s As Long, si As Long, sj As Long, mi As Long
    Dim ox As Long, sx As Long, tx As Long
    Dim i As Long, j As Long
    Dim k As Long
    
    ' --- Sheet 相关 ---
    Dim sheetCount As Long, sheetCount2 As Long
    Dim mSheet As Long, lastSh As Long, sheetId As Long
    Dim shNm() As Object
    Dim nd As Object
    Dim sp As paths
    
    ' --- Operation / SubOperation ---
    Dim opN As Operation, opM As Operation, oc As Operation
    Dim sbN As SubOperations, sbM As SubOperations, sc As SubOperations
    Dim subF As SubOperation, subN As SubOperation
    Dim subL As SubOperation, subM As SubOperation, sbc As SubOperation
    
    ' --- ToolPath ---
    Dim tpF As paths, tpN As paths, tpL As paths, tpsM As paths
    Dim tpc As paths, tpc2 As Path
    Dim tF As Path, tN2 As Path, tL As Path, tpM As Path
    Dim ta As Path
    
    ' --- Tool ---
    Dim tM As MillTool
    
    ' --- 字典/集合 ---
    Dim stD As Object
    Dim nc As Collection
    Dim c2 As Collection
    Dim tc As Collection
    
    ' --- 字符串/数值 ---
    Dim firstTpName As String, lookupName As String
    Dim mn As String, tk As String, ck As String
    Dim tky As String, ck2 As String
    Dim currentSig As String
    Dim pos As Long
    Dim spInt As Integer
    Dim siN As Long
    
    ' ==================================================================
    ' 函数体开始
    ' ==================================================================
    
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    
    Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    
    ' 检测图纸是否已变更（用户可能在非模态对话框打开时切换了图纸）
    currentSig = drw.Name & "|" & CStr(ops.Count)
    If currentSig <> g_lastDrawingSig Then
        Set g_mapPathToSheet = Nothing        ' 图纸已变，清空缓存强制重建
        g_lastDrawingSig = currentSig
    End If
    
    ' ==================================================================
    ' 阶段1：建立 g_mapPathToSheet — 每个 Operation 第一条刀具路径 → Sheet 编号
    ' ==================================================================
    If g_mapPathToSheet Is Nothing Then
        Set g_mapPathToSheet = CreateObject("Scripting.Dictionary")
        
        ' 读取 NestInformation（排版信息），获取 Sheet 列表
        Set ni = drw.GetNestInformation()
        If ni Is Nothing Then
            ' 无 NestInformation 时视为单一 Sheet
            sheetCount = 1
        Else
            sheetCount = ni.Sheets.Count
            If sheetCount = 0 Then sheetCount = 1   ' 无排版时视为单一 Sheet
        End If
        
        ' shNm(s) 是第 s 个 Sheet 上所有路径名的字典（用于快速查找）
        ReDim shNm(1 To sheetCount)
        
        ' 遍历每个 Sheet，收集其包含的所有路径名
        For s = 1 To sheetCount
            Set nd = CreateObject("Scripting.Dictionary")
            
            ' ni.Sheets(s) 在 sheetCount 范围内一定存在（因为上面已处理 Count=0 的情况）
            If ni Is Nothing Then
                ' 无 NestInformation，shNm 留空字典
            ElseIf s <= ni.Sheets.Count Then
                Set sp = ni.Sheets(s).paths            ' Sheet 上的所有路径
                
                For mi = 1 To sp.Count
                    If Not nd.Exists(sp(mi).Name) Then nd.Add sp(mi).Name, True
                Next mi
            End If
            
            Set shNm(s) = nd                        ' 本 Sheet 的路径名集合
        Next s
        
        ' 遍历所有 Operation，判断其第一条刀具路径属于哪个 Sheet
        For opIdx = 1 To ops.Count
            Set opN = ops(opIdx)
            
            Set sbN = opN.SubOperations
            
            mSheet = 0
            firstTpName = ""                        ' 本 Operation 的第一条路径名
            
            ' 获取本 Operation 第一条 SubOperation 的第一条路径名
            If Not (sbN Is Nothing) Then
                If sbN.Count > 0 Then
                    Set subF = sbN(1)
                    
                    Set tpF = subF.ToolPaths
                    
                    If Not (tpF Is Nothing) Then
                        If tpF.Count > 0 Then
                            Set tF = tpF(1)
                            If Not (tF Is Nothing) Then firstTpName = tF.Name
                        End If
                    End If
                End If
            End If
            
            ' 在所有 Sheet 的路径集合中查找该路径名，确定所在 Sheet
            If Not (sbN Is Nothing) Then
                For siN = 1 To sbN.Count
                    Set subN = sbN(siN)
                    
                    Set tpN = subN.ToolPaths
                    
                    If Not (tpN Is Nothing) Then
                        For mi = 1 To tpN.Count
                            Set tN2 = tpN(mi)
                            
                            If Not (tN2 Is Nothing) Then
                                ' 在所有 Sheet 中查找此路径名
                                For s = 1 To sheetCount
                                    If shNm(s).Exists(tN2.Name) Then
                                        mSheet = s
                                        lastSh = s          ' 记录最后一个有匹配的 Sheet
                                        Exit For
                                    End If
                                Next s
                                If mSheet > 0 Then Exit For   ' 已找到，跳出路径循环
                            End If
                        Next mi
                    End If
                    If mSheet > 0 Then Exit For               ' 已找到，跳出 SubOp 循环
                Next siN
            End If
            
            ' 如果完全没找到，用上一个匹配的 Sheet 作为兜底
            If mSheet = 0 And lastSh > 0 Then mSheet = lastSh
            If mSheet = 0 Then mSheet = 1                     ' 仍未找到，默认 Sheet 1
            
            ' 缓存：路径名 → Sheet 编号
            If firstTpName <> "" Then
                If Not g_mapPathToSheet.Exists(firstTpName) Then
                    g_mapPathToSheet.Add firstTpName, mSheet
                End If
            End If
        Next opIdx
    End If
    
    ' ==================================================================
    ' 阶段2：按 "SheetID|加工方式+刀具" 分组，收集所有刀具路径
    ' ==================================================================
    Set stD = CreateObject("Scripting.Dictionary")
    
    For opIdx = 1 To ops.Count
        Set opM = ops(opIdx)
        
        Set sbM = opM.SubOperations
        If Not (sbM Is Nothing) Then
            
            ' 先获取本 Operation 的第一条路径名，用于查 SheetId
            lookupName = ""
            If sbM.Count > 0 Then
                Set subL = sbM(1)
                
                Set tpL = subL.ToolPaths
                
                If Not (tpL Is Nothing) Then
                    If tpL.Count > 0 Then
                        Set tL = tpL(1)
                        If Not (tL Is Nothing) Then lookupName = tL.Name
                    End If
                End If
            End If
            
            ' 查 SheetId
            sheetId = 1
            If lookupName <> "" And g_mapPathToSheet.Exists(lookupName) Then
                sheetId = g_mapPathToSheet(lookupName)
            End If
            
            ' 遍历本 Operation 所有 SubOperation
            For si = 1 To sbM.Count
                Set subM = sbM(si)
                
                Set tM = subM.Tool
                If Not (tM Is Nothing) Then
                    ' 提取加工方式名（去尾）
                    mn = subM.Name
                    spInt = InStr(mn, "  ")
                    If spInt > 0 Then
                        mn = Left(mn, spInt - 1)
                    Else
                        spInt = InStr(mn, " ")
                        If spInt > 0 Then mn = Left(mn, spInt - 1)
                    End If
                    
                    ' 构造分组键（无 Sheet 前缀的原始键，用于匹配 sortedKeys）
                    tk = mn & " T" & CStr(tM.Number) & " " & tM.Name
                    
                    ' 构造分 Sheet 分组键（带 Sheet 前缀）
                    ck = CStr(sheetId) & "|" & tk
                    
                    ' 收集本 SubOperation 的所有刀具路径
                    Set tpsM = subM.ToolPaths
                    
                    If Not (tpsM Is Nothing) Then
                        For mi = 1 To tpsM.Count
                            Set tpM = tpsM(mi)
                            
                            If Not (tpM Is Nothing) Then
                                ' 首次遇到此分组键时创建 Collection
                                If Not stD.Exists(ck) Then
                                    Set nc = New Collection
                                    stD.Add ck, nc
                                End If
                                
                                Set c2 = stD(ck)
                                c2.Add tpM                              ' 收集路径引用
                            End If
                        Next mi
                    End If
                End If
            Next si
        End If
    Next opIdx
    
    If stD.Count = 0 Then Exit Sub                        ' 没有可排序的内容
    
    ' ==================================================================
    ' 阶段3：清空旧 OpNo → 重新编号 → OrderAll 排序
    ' ==================================================================
    
    ' --- 3a: 将所有刀具路径的 OpNo 清零 ---
    For ox = 1 To ops.Count
        Set oc = ops(ox)
        
        Set sc = oc.SubOperations
        If Not (sc Is Nothing) Then
            For sx = 1 To sc.Count
                Set sbc = sc(sx)
                
                Set tpc = sbc.ToolPaths
                If Not (tpc Is Nothing) Then
                    For tx = 1 To tpc.Count
                        Set tpc2 = tpc(tx)
                        If Not (tpc2 Is Nothing) Then tpc2.OpNo = 0       ' 清零
                    Next tx
                End If
            Next sx
        End If
    Next ox
    
    ' --- 3b: 设置撤销点 ---
    App.SetUndoCommandName "排版刀具排序"
    App.SetUndoPoint
    
    ' --- 3c: 禁用屏幕刷新，批量操作 ---
    drw.ScreenUpdating = False
    
    ' --- 3d: 按 Sheet 遍历，为每个分组分配 OpNo ---
    ' OpNo = sheetIndex × 100 + pos
    ' 每个 Sheet 从 101, 102, 103... 开始编号
    ' 这样 OrderAll 会先按 Sheet 分组（百位），再按组内顺序（个位）
    sheetCount2 = 1
    
    ' 重新获取 sheetCount（可能在阶段1缓存未命中时未初始化）
    Set ni2 = drw.GetNestInformation()
    If Not (ni2 Is Nothing) Then
        sheetCount2 = ni2.Sheets.Count
        If sheetCount2 = 0 Then sheetCount2 = 1
    End If
    
    For si = 1 To sheetCount2                                  ' 按 Sheet 顺序
        pos = 1                                                 ' 组内序号从 1 开始
        For sj = 0 To UBound(sortedKeys)                        ' 按用户指定顺序
            tky = sortedKeys(sj)
            
            ' 构造本 Sheet 的分组键
            ck2 = CStr(si) & "|" & tky
            
            If stD.Exists(ck2) Then
                Set tc = stD(ck2)                                ' 取出本组所有路径
                
                For mi = 1 To tc.Count
                    Set ta = tc(mi)
                    If Not (ta Is Nothing) Then
                        ta.OpNo = si * 100 + pos                 ' 设置排序编号
                    End If
                Next mi
                pos = pos + 1                                    ' 下一组序号
            End If
        Next sj
    Next si
    
    ' --- 3e: 让 AlphaCAM 按 OpNo 重新排序所有刀具路径 ---
    ops.OrderAll
    
    ' --- 3f: 恢复屏幕刷新并重绘 ---
    drw.ScreenUpdating = True
    drw.Redraw
    
    Exit Sub
    
' ==============================================================================
' 错误处理：确保屏幕刷新恢复
' ==============================================================================
ErrHandler3:
    If Not (drw Is Nothing) Then
        drw.ScreenUpdating = True
    End If
    MsgBox "排序出错: " & Err.Description, vbCritical, "排版刀具排序"
End Sub
