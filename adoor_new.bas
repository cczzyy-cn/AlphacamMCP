' C梯形带内边
Public Sub AdoorMain(RequiredData As Object, _
                     Optional User_1 As Variant, Optional User_2 As Variant, _
                     Optional User_3 As Variant, Optional User_4 As Variant, _
                     Optional User_5 As Variant, Optional User_6 As Variant, _
                     Optional User_7 As Variant)
   Dim UserDims As Collection
   Set UserDims = RequiredData.UserVariables
   On Error GoTo AdoorMain_Error

' Main Variables
   Dim width As Double
   Dim length As Double
   Dim corner_radius As Double
   width = RequiredData.width
   length = RequiredData.length
   corner_radius = RequiredData.CornerRadius

' User Variables
   Dim H2 As Double
   Dim W As Double
   Dim L0R1 As Double
   Dim X As Double
   Dim B As Double
   H2 = UserDims(1)
   W = UserDims(2)
   L0R1 = UserDims(3)
   B = UserDims(4)
   X = L0R1 * (width - (W * 2))

   Dim FastGeo1 As FastGeometry
   Dim Geo1 As Path
   Set FastGeo1 = App.ActiveDrawing.CreateFastGeometry
   With FastGeo1
      .Point width * L0R1, 0
      .Point width * L0R1, H2
      .Point width - W + L0R1 * (width - (width - W) * 2), length
      .Point width - L0R1 * width, length
      .Point width - L0R1 * width, 0
      Set Geo1 = .CloseAndFinish
   End With
   Geo1.Group = 1
   RequiredData.PathsToReturn.Add Geo1

' 图形2 = 图形1 向内偏移 B（偏移侧随 L0R1 镜像的路径方向切换：
' L0R1=0 路径为 CW，内偏移用 Right(-1)；L0R1=1 路径为镜像 CCW，内偏移用 Left(1)）
   Dim Geo2 As Path
   Dim Offs As Object
   Dim OffsSide As Integer
   If L0R1 = 0 Then OffsSide = -1 Else OffsSide = 1
   Set Offs = Geo1.Offset(B, OffsSide)
   If Not Offs Is Nothing Then
      If Offs.Count > 0 Then
         Set Geo2 = Offs.Item(1)
         Geo2.Group = 2
         RequiredData.PathsToReturn.Add Geo2
      End If
   End If

Controlled_Exit:
   With RequiredData
      If .PathsToReturn.Count = 0 Then Set .PathsToReturn = Nothing
   End With
   Exit Sub

AdoorMain_Error:
   MsgBox Err.Description, vbExclamation, Err.Source
   Set RequiredData.PathsToReturn = Nothing
   Resume Controlled_Exit

End Sub

Public Function Sindeg(AngleInDegrees As Double) As Double
   Sindeg = Sin(AngleInDegrees * (4 * Atn(1)) / 180)
End Function

Public Function Cosdeg(AngleInDegrees As Double) As Double
   Cosdeg = Cos(AngleInDegrees * (4 * Atn(1)) / 180)
End Function

Public Function Tandeg(AngleInDegrees As Double) As Double
   Tandeg = Tan(AngleInDegrees * (4 * Atn(1)) / 180)
End Function

Public Function InvSin(X As Double) As Double
   InvSin = Atn(X / Sqr(-X * X + 1)) * (180 / (4 * Atn(1)))
End Function

Public Function InvCos(X As Double) As Double
   InvCos = (Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)) * (180 / (4 * Atn(1)))
End Function

Public Function InvTan(X As Double) As Double
   InvTan = Atn(X) * (180 / (4 * Atn(1)))
End Function





