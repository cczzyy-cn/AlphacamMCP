# -*- coding: utf-8 -*-
"""Update AdoorEvents module: mirror door profile about vertical centre line
(X = width/2) when L0orR1 = 0."""
import win32com.client

NEW_CODE = """Public Sub AdoorMain(RequiredData As Object, _
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
   Dim H As Double
   Dim W As Double
   Dim R1 As Double
   Dim R2 As Double
   Dim L0orR1 As Double
   H = UserDims(1)
   W = UserDims(2)
   R1 = UserDims(3)
   R2 = UserDims(4)
   L0orR1 = UserDims(5)

' L0orR1 = 0 -> mirror the profile about the vertical centre line (X = width / 2)
   Dim mirror As Boolean
   mirror = (L0orR1 = 0)

   Dim FastGeo1 As FastGeometry
   Dim Geo1 As Path
   Set FastGeo1 = App.ActiveDrawing.CreateFastGeometry
   With FastGeo1
      .Point MirrorX(0, width, mirror), length
      .Point MirrorX(0, width, mirror), 0
      .Point MirrorX(W, width, mirror), 0
      .Point MirrorX(W, width, mirror), length - H - R2
      .KnownArc R2, Not mirror, MirrorX(W + R2, width, mirror), length - H - R2
      .Point MirrorX(W + R2, width, mirror), length - H
      .Point MirrorX(width - R1, width, mirror), length - H
      .KnownArc R1, mirror, MirrorX(width - R1, width, mirror), length - H + R1
      .Point MirrorX(width, width, mirror), length - H + R1
      .Point MirrorX(width, width, mirror), length
      Set Geo1 = .CloseAndFinish
   End With
   RequiredData.PathsToReturn.Add Geo1

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

' MirrorX: map X about the vertical centre line (X = width / 2) when mirror
' is True, otherwise return the coordinate unchanged.
Private Function MirrorX(ByVal X As Double, ByVal width As Double, ByVal mirror As Boolean) As Double
   If mirror Then
      MirrorX = width - X
   Else
      MirrorX = X
   End If
End Function

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
"""

app = win32com.client.GetActiveObject('aroutaps.Application')
vbe = app.VBE
proj = vbe.ActiveVBProject
comp = None
for j in range(1, proj.VBComponents.Count + 1):
    c = proj.VBComponents(j)
    if c.Name == 'AdoorEvents':
        comp = c
        break
assert comp is not None, 'AdoorEvents module not found'
cm = comp.CodeModule

old = cm.Lines(1, cm.CountOfLines)
assert 'Public Sub AdoorMain' in old, 'AdoorMain not found in module'
assert '.KnownArc R2, True, W + R2' in old, 'expected old arc pattern not found'

cm.DeleteLines(1, cm.CountOfLines)
cm.AddFromString(NEW_CODE)

new = cm.Lines(1, cm.CountOfLines)
checks = {
    'lines': cm.CountOfLines,
    'has MirrorX func': 'Private Function MirrorX' in new,
    'has mirror flag': 'mirror = (L0orR1 = 0)' in new,
    'R2 arc mirrored CW': '.KnownArc R2, Not mirror' in new,
    'R1 arc mirrored CW': '.KnownArc R1, mirror' in new,
    'old pattern removed': '.KnownArc R2, True, W + R2' not in new,
    'has Sindeg': 'Public Function Sindeg' in new,
    'has InvTan': 'Public Function InvTan' in new,
}
for k, v in checks.items():
    print(k, ':', v)
print('--- first 40 lines ---')
print('\n'.join(new.splitlines()[:40]))
