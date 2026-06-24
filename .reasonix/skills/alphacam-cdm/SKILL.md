---
name: alphacam-cdm
description: Generate CDM door panel template VBA code (AdoorMain callback, RequiredData, Geo2D, Fillet)
---

# AlphaCAM CDM 门板模板开发

你是一个 AlphaCAM CDM (Custom Door Maker) 门板模板专家。根据用户的需求生成正确的 CDM 门板模板 VBA 代码。

## 入口点

```vb
Public Sub AdoorMain(RequiredData As Object, _
                     User_1 As Variant, User_2 As Variant, _
                     User_3 As Variant, User_4 As Variant, _
                     User_5 As Variant, User_6 As Variant, _
                     User_7 As Variant)
```

必须加 `On Error GoTo ErrorHandler` 和错误处理标签。

## RequiredData 对象

| 成员 | 类型 | 说明 |
|------|------|------|
| `.Width` | Double | 面板宽度 |
| `.Length` | Double | 面板高度 |
| `.CornerRadius` | Double | 面板圆角半径 |
| `.UserVariables` | Collection | 用户定义尺寸（用 Enum 索引） |
| `.PathsToReturn.Add(path)` | Method | 返回创建的几何 |
| `.Success` | Boolean | 设为 False 标记错误 |

## 几何创建流程

```vb
' 1. 用 Enum 定义用户尺寸索引
Private Enum AdoorStyleDimensions
    adoorDimSTILE = 1
    adoorDimRAIL = 2
    ' ...
End Enum

' 2. 获取尺寸
dblWidth = RequiredData.Width
dblStile = UserDims(adoorDimSTILE)

' 3. 创建几何
Set geoIn = App.ActiveDrawing.Create2DGeometry(startX, startY)
With geoIn
    .AddLine x, y           ' 画直线到 (x, y)
    .AddArc2Point mx,my, ex,ey, angle?  ' 三点画弧（参数待确认）
    Set pthIn = .CloseAndFinishLine     ' 闭合返回 Path
End With

' 4. 倒圆角/设置属性
With pthIn
    .Group = 1
    .ToolInOut = acamINSIDE
    .CW = False
    .Fillet radius
End With

' 5. 返回几何
RequiredData.PathsToReturn.Add pthIn
```

## 错误处理模板

```vb
On Error GoTo ErrorHandler

' ... 主代码 ...

Controlled_Exit:
    Set pthIn = Nothing
    Set geoIn = Nothing
    With RequiredData
        If .PathsToReturn.Count = 0 Then Set .PathsToReturn = Nothing
    End With
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbExclamation, Err.Source
    Set RequiredData.PathsToReturn = Nothing
    RequiredData.Success = False
    Resume Controlled_Exit
```

## 长门模式（下矩形）

当需要长门样式时（下矩形高度 > 0）：

1. **下矩形**从底部内框线（`dblRailBottom`）**向上**画，宽度在左右边框之间
2. **上方弧形图形底部**要相应**上移**，给下矩形让出空间
3. 上下图形之间的间距由参数控制

```vb
' 计算上方图形底部位置
Dim dblUpperBottom As Double
dblUpperBottom = dblRailBottom
If dblLowerRectHeight > 0 Then
    dblUpperBottom = dblRailBottom + dblLowerRectHeight + dblMiddleSpacing
End If

' 上方图形（从 dblUpperBottom 开始）
Set geoIn = App.ActiveDrawing.Create2DGeometry(dblStileLeft, dblUpperBottom)
' ... 画弧形 ...
.AddLine (dblWidth - dblStile), dblUpperBottom   ' 底部线用 dblUpperBottom
Set pthIn = .CloseAndFinishLine

' 下方矩形
If dblLowerRectHeight > 0 Then
    Set geoLower = App.ActiveDrawing.Create2DGeometry(dblStileLeft, dblRailBottom)
    .AddLine dblStileLeft, (dblRailBottom + dblLowerRectHeight)
    .AddLine (dblWidth - dblStile), (dblRailBottom + dblLowerRectHeight)
    .AddLine (dblWidth - dblStile), dblRailBottom
    Set pthLower = .CloseAndFinishLine
    RequiredData.PathsToReturn.Add pthLower
End If
```

## 内偏移（第三个图形）

用 `Path.Offset(Distance, Side)` 对已有路径做内偏移：

```vb
If dblInnerOffset > 0 Then
    Dim offsetPaths As Paths
    ' 逆时针时 LEFT = 向内
    Set offsetPaths = pthIn.Offset(dblInnerOffset, 1)  ' 1 = acamLEFT
    If Not (offsetPaths Is Nothing) Then
        For o = 1 To offsetPaths.Count
            Set pthOffset = offsetPaths(o)
            pthOffset.Group = 3
            RequiredData.PathsToReturn.Add pthOffset
        Next o
    End If
End If
```

## 分组规则

| Group | 图形 |
|-------|------|
| 1 | 上方圆弧图形 |
| 2 | 下方矩形（长门模式） |
| 3 | 内偏移图形 |
