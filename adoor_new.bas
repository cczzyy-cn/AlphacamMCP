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
   Dim X As Double
   Dim y As Double
   Dim h As Double
   Dim r As Double
   X = UserDims(1)
   y = UserDims(2)
   h = UserDims(3)
   r = h / 4 + (width - X * 2) ^ 2 / (16 * h)
   Dim GroupNumber1 As Integer
   Dim GroupNumber2 As Integer
   GroupNumber1 = App.ActiveDrawing.GetNextGroupNumberForGeometries
   GroupNumber2 = GroupNumber1 + 1
   Dim FastGeo1 As FastGeometry
   Dim Geo1 As Path
   Set FastGeo1 = App.ActiveDrawing.CreateFastGeometry
   With FastGeo1
      .Point X, y
      .Point width - X, y
      .Point width - X, length - y - h
      .KnownArc r, True, width - X, length - y - h + r
      .KnownArc r, False, width / 2, length - y - r
      .KnownArc r, True, X, length - y - h + r
      .Point X, length - y - h
      .Point X, y
      Set Geo1 = .Finish
   End With
   Geo1.Group = GroupNumber1
   RequiredData.PathsToReturn.Add Geo1
   Dim FastGeo2 As FastGeometry
   Dim Geo2 As Path
   Set FastGeo2 = App.ActiveDrawing.CreateFastGeometry
   With FastGeo2
      .Point X, length
      .Point X, 0
      Set Geo2 = .Finish
   End With
   Geo2.Group = GroupNumber2
   RequiredData.PathsToReturn.Add Geo2
   Dim FastGeo3 As FastGeometry
   Dim Geo3 As Path
   Set FastGeo3 = App.ActiveDrawing.CreateFastGeometry
   With FastGeo3
      .Point width - X, 0
      .Point width - X, length
      Set Geo3 = .Finish
   End With
   Geo3.Group = GroupNumber2
   RequiredData.PathsToReturn.Add Geo3
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
