Imports System
Imports System.Collections.Generic
Imports AlphaCAMRouter

Namespace 小条先切
	' Token: 0x0200000E RID: 14
	Public Class OpsPara
		' Token: 0x17000030 RID: 48
		' (get) Token: 0x06000096 RID: 150 RVA: 0x000026C9 File Offset: 0x000008C9
		' (set) Token: 0x06000097 RID: 151 RVA: 0x000026D1 File Offset: 0x000008D1
		Public Property Boolean_0 As Boolean
			Get
				Return Me._include3d
			End Get
			Set(value As Boolean)
				Me._include3d = value
			End Set
		End Property

		' Token: 0x17000031 RID: 49
		' (get) Token: 0x06000098 RID: 152 RVA: 0x000026DA File Offset: 0x000008DA
		' (set) Token: 0x06000099 RID: 153 RVA: 0x000026E2 File Offset: 0x000008E2
		Public Property IncludeOpen As Boolean
			Get
				Return Me._includeOpen
			End Get
			Set(value As Boolean)
				Me._includeOpen = value
			End Set
		End Property

		' Token: 0x17000032 RID: 50
		' (get) Token: 0x0600009A RID: 154 RVA: 0x000026EB File Offset: 0x000008EB
		' (set) Token: 0x0600009B RID: 155 RVA: 0x000026F3 File Offset: 0x000008F3
		Public Property Op As Operation
			Get
				Return Me._op
			End Get
			Set(value As Operation)
				Me._op = value
			End Set
		End Property

		' Token: 0x17000033 RID: 51
		' (get) Token: 0x0600009C RID: 156 RVA: 0x000026FC File Offset: 0x000008FC
		' (set) Token: 0x0600009D RID: 157 RVA: 0x00002704 File Offset: 0x00000904
		Public Property ToolName As String
			Get
				Return Me._toolName
			End Get
			Set(value As String)
				Me._toolName = value
			End Set
		End Property

		' Token: 0x17000034 RID: 52
		' (get) Token: 0x0600009E RID: 158 RVA: 0x0000270D File Offset: 0x0000090D
		' (set) Token: 0x0600009F RID: 159 RVA: 0x00002715 File Offset: 0x00000915
		Public Property Tps As List(Of CutToolPath)
			Get
				Return Me._tps
			End Get
			Set(value As List(Of CutToolPath))
				Me._tps = value
			End Set
		End Property

		' Token: 0x060000A0 RID: 160 RVA: 0x0000271E File Offset: 0x0000091E
		Public Sub New(bool_0 As Boolean, includeOpen As Boolean, toolName As String, op As Operation)
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
			Me.Op = op
			Me.Boolean_0 = bool_0
			Me.IncludeOpen = includeOpen
			Me.ToolName = toolName
			Me.Tps = New List(Of CutToolPath)()
		End Sub

		' Token: 0x060000A1 RID: 161 RVA: 0x00006660 File Offset: 0x00004860
		Public Sub [Set]()
			Dim subOperations As SubOperations = Me.Op.SubOperations
			Dim count As Integer = subOperations.Count
			For i As Integer = 1 To count
				Dim subOperation As SubOperation = subOperations.Item(i)
				If String.Compare(subOperation.Name, Me.ToolName) = 0 Then
					Dim toolPaths As Paths = subOperation.ToolPaths
					Dim count2 As Integer = toolPaths.Count
					For j As Integer = 1 To count2
						Dim path As Path = toolPaths.Item(j)
						If(Not path.Is3D OrElse Me.Boolean_0) AndAlso (path.Closed OrElse Me.IncludeOpen) Then
							Dim cutToolPath As CutToolPath = New CutToolPath()
							cutToolPath.ToolPath = path
							cutToolPath.MillData = path.GetMillData()
							cutToolPath.FinalDepth = CDbl(cutToolPath.MillData.FinalDepth)
							cutToolPath.DepthCutSpecial = If(cutToolPath.MillData.DepthsOfCutSpecified, 1, 0)
							Dim num As Double
							Dim num2 As Double
							Dim num3 As Double
							Dim num4 As Double
							cutToolPath.ToolPath.GetFeedExtent(num, num2, num3, num4)
							cutToolPath.MinX = num
							cutToolPath.MaxX = num3
							cutToolPath.MinY = num2
							cutToolPath.MaxY = num4
							cutToolPath.CW = cutToolPath.ToolPath.CW = 1S
							cutToolPath.IsDelete = False
							cutToolPath.Horizontal = cutToolPath.MaxX - cutToolPath.MinX >= cutToolPath.MaxY - cutToolPath.MinY
							cutToolPath.ToolFileName = subOperation.Tool.FileName
							Me.Tps.Add(cutToolPath)
						End If
					Next
				End If
			Next
		End Sub

		' Token: 0x04000043 RID: 67
		Private _include3d As Boolean

		' Token: 0x04000044 RID: 68
		Private _includeOpen As Boolean

		' Token: 0x04000045 RID: 69
		Private _op As Operation

		' Token: 0x04000046 RID: 70
		Private _toolName As String

		' Token: 0x04000047 RID: 71
		Private _tps As List(Of CutToolPath)
	End Class
End Namespace
