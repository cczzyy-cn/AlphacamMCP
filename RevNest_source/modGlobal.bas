in file: C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\000\StartUp\Utils\ReverseNest\ReverseNest.amb - OLE stream: 'vao/The VBA Project/_VBA_Project/VBA/modGlobal'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

' >< ADD-IN SPECIFIC CONSTANTS ><
'
' !! the following MUST be altered accordingly
'
' ->
'
Public Const DEF_APP_TITLE                  As String = "Reverse Side Nesting"
Public Const DEF_VERSION                    As String = "1.2"
Public Const DEF_MACRO_NAME                 As String = "ReverseNest"
Public Const DEF_CTX                        As String = "ReverseNest.txt"
'
' <-

' >< API ><
'
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

' >< ENUM ><
'
Public Enum AlphaIntersectPoint
        alphaIntersect_PARALLEL = -1
        alphaIntersect_NONE = 0
        alphaIntersect_LINE_1 = 1
        alphaIntersect_LINE_2 = 2
        alphaIntersect_BOTH_LINES = 3
End Enum
        
' >< UDT ><
'
Public Type POINT_XYZ
        X                           As Double
        Y                           As Double
        Z                           As Double
End Type

Public Type WP_XYZ
        X                           As POINT_XYZ
        Y                           As POINT_XYZ
        Z                           As POINT_XYZ
        Origin                      As POINT_XYZ
End Type

Public Type LINE_XYZ
        StartPoint                  As POINT_XYZ    'Starting point (X,Y,Z) on line.
        EndPoint                    As POINT_XYZ    'Ending point (X,Y,Z) on line.
End Type

Public Type ARC_DETAILS
        IsValidArc                  As Boolean      'Is this a valid arc.
        StartPoint                  As POINT_XYZ    'Starting point.
        MidPoint                    As POINT_XYZ    'Mid point.
        EndPoint                    As POINT_XYZ    'Ending point.
        CenterPoint                 As POINT_XYZ    'Center point.
        Radius                      As Double       'Radius.
        StartAngle                  As Double       'Starting angle in radians.
        MidAngle                    As Double       'Mid angle in radians.
        EndAngle                    As Double       'Ending angle in radians.
End Type

' >< CONSTANTS ><
'
Public Const Pi                     As Double = 3.14159265358979    'Pi
'
Public Const LicomUKsab_nest        As String = "LicomUKsab_nest"
Public Const LicomUKsab_path        As String = "LicomUKsab_path"
Public Const LicomUKsab_part        As String = "LicomUKsab_part"
Public Const LicomUKsab_outer       As String = "LicomUKsab_outer"
Public Const LicomUKsab_sheet       As String = "LicomUKsab_sheet"
Public Const LicomUKja_part         As String = "LicomUKja_part"
Public Const LicomUKjba_part        As String = "LicomUKjba_part"


-------------------------------------------------------------------------------
