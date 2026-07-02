Imports System
Imports System.ComponentModel

Namespace 小条先切
	' Token: 0x02000003 RID: 3
	Public Class [Set]
		Implements INotifyPropertyChanged

		' Token: 0x17000001 RID: 1
		' (get) Token: 0x06000001 RID: 1 RVA: 0x0000223C File Offset: 0x0000043C
		' (set) Token: 0x06000002 RID: 2 RVA: 0x00002244 File Offset: 0x00000444
		Public Property ToolName As String
			Get
				Return Me._toolName
			End Get
			Set(value As String)
				Me._toolName = value
				If Me.PropertyChanged IsNot Nothing Then
					Me.PropertyChanged(Me, New PropertyChangedEventArgs("ToolName"))
				End If
			End Set
		End Property

		' Token: 0x17000002 RID: 2
		' (get) Token: 0x06000003 RID: 3 RVA: 0x0000226B File Offset: 0x0000046B
		' (set) Token: 0x06000004 RID: 4 RVA: 0x00002273 File Offset: 0x00000473
		Public Property SmallWidth As Double
			Get
				Return Me._smallWidth
			End Get
			Set(value As Double)
				Me._smallWidth = value
				If Me.PropertyChanged IsNot Nothing Then
					Me.PropertyChanged(Me, New PropertyChangedEventArgs("SmallWidth"))
				End If
			End Set
		End Property

		' Token: 0x17000003 RID: 3
		' (get) Token: 0x06000005 RID: 5 RVA: 0x0000229A File Offset: 0x0000049A
		' (set) Token: 0x06000006 RID: 6 RVA: 0x000022A2 File Offset: 0x000004A2
		Public Property IsSloopingLine As Boolean
			Get
				Return Me._isSloopingLine
			End Get
			Set(value As Boolean)
				Me._isSloopingLine = value
				If Me.PropertyChanged IsNot Nothing Then
					Me.PropertyChanged(Me, New PropertyChangedEventArgs("IsSloopingLine"))
				End If
			End Set
		End Property

		' Token: 0x17000004 RID: 4
		' (get) Token: 0x06000007 RID: 7 RVA: 0x000022C9 File Offset: 0x000004C9
		' (set) Token: 0x06000008 RID: 8 RVA: 0x000022D1 File Offset: 0x000004D1
		Public Property SloopDist As Double
			Get
				Return Me._sloopDist
			End Get
			Set(value As Double)
				Me._sloopDist = value
				If Me.PropertyChanged IsNot Nothing Then
					Me.PropertyChanged(Me, New PropertyChangedEventArgs("IsSloopingLine"))
				End If
			End Set
		End Property

		' Token: 0x17000005 RID: 5
		' (get) Token: 0x06000009 RID: 9 RVA: 0x000022F8 File Offset: 0x000004F8
		' (set) Token: 0x0600000A RID: 10 RVA: 0x000028C0 File Offset: 0x00000AC0
		Public Property BackDist As Double
			Get
				Return Me._backDist
			End Get
			Set(value As Double)
				If value < 0.0 Then
					Me._backDist = 0.0
				Else
					Me._backDist = value
				End If
				If Me.PropertyChanged IsNot Nothing Then
					Me.PropertyChanged(Me, New PropertyChangedEventArgs("BackDist"))
				End If
			End Set
		End Property

		' Token: 0x17000006 RID: 6
		' (get) Token: 0x0600000B RID: 11 RVA: 0x00002300 File Offset: 0x00000500
		' (set) Token: 0x0600000C RID: 12 RVA: 0x00002308 File Offset: 0x00000508
		Public Property IsUniqueCutTool As Boolean
			Get
				Return Me._isUniqueCutTool
			End Get
			Set(value As Boolean)
				Me._isUniqueCutTool = value
				If Me.PropertyChanged IsNot Nothing Then
					Me.PropertyChanged(Me, New PropertyChangedEventArgs("IsUniqueCutTool"))
				End If
			End Set
		End Property

		' Token: 0x0600000D RID: 13 RVA: 0x00002910 File Offset: 0x00000B10
		Public Sub New()
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
			Me.ToolName = ""
			Me.IsSloopingLine = True
			Me.SloopDist = 20.0
			Me.SmallWidth = 260.0
			Me.BackDist = 40.0
			Me.IsUniqueCutTool = False
		End Sub

		' Token: 0x0600000E RID: 14 RVA: 0x00002970 File Offset: 0x00000B70
		Public Sub Copy(obj As [Set])
			Me.ToolName = obj.ToolName
			Me.SmallWidth = obj.SmallWidth
			Me.IsSloopingLine = obj.IsSloopingLine
			Me.SloopDist = obj.SloopDist
			Me.BackDist = obj.BackDist
			Me.IsUniqueCutTool = obj.IsUniqueCutTool
		End Sub

		' Token: 0x0600000F RID: 15 RVA: 0x000029C8 File Offset: 0x00000BC8
		Public Function Equal(obj As [Set]) As Boolean
			Return String.Compare(Me.ToolName, obj.ToolName) = 0 AndAlso Me.SmallWidth = obj.SmallWidth AndAlso Me.IsSloopingLine = obj.IsSloopingLine AndAlso Me.SloopDist = obj.SloopDist AndAlso Me.BackDist = obj.BackDist AndAlso Me.IsUniqueCutTool = obj.IsUniqueCutTool
		End Function

		' Token: 0x06000010 RID: 16 RVA: 0x00002A3C File Offset: 0x00000C3C
		Public Function ToString() As String
			Return String.Concat(New Object() { Me.ToolName, "|", Me.SmallWidth, "|", Me.IsSloopingLine, "|", Me.SloopDist, "|", Me.IsUniqueCutTool })
		End Function

		' Token: 0x14000001 RID: 1
		' (add) Token: 0x06000011 RID: 17 RVA: 0x00002AB8 File Offset: 0x00000CB8
		' (remove) Token: 0x06000012 RID: 18 RVA: 0x00002AF0 File Offset: 0x00000CF0
		Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

		' Token: 0x04000001 RID: 1
		Private _toolName As String

		' Token: 0x04000002 RID: 2
		Private _smallWidth As Double

		' Token: 0x04000003 RID: 3
		Private _isSloopingLine As Boolean

		' Token: 0x04000004 RID: 4
		Private _sloopDist As Double

		' Token: 0x04000005 RID: 5
		Private _backDist As Double

		' Token: 0x04000006 RID: 6
		Private _isUniqueCutTool As Boolean
	End Class
End Namespace
