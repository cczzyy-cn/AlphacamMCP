Option Explicit

Public Width As Double
Public Length As Double
Public CornerRadius As Double
Public UserVariables As Collection
Public PathsToReturn As Collection
Public Success As Boolean

Private Sub Class_Initialize()
    Set UserVariables = New Collection
    Set PathsToReturn = New Collection
End Sub
