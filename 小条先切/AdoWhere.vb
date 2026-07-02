Imports System

Namespace 小条先切
	' Token: 0x02000009 RID: 9
	Public Class AdoWhere
		' Token: 0x1700002C RID: 44
		' (get) Token: 0x0600007B RID: 123 RVA: 0x00002635 File Offset: 0x00000835
		' (set) Token: 0x0600007C RID: 124 RVA: 0x0000263D File Offset: 0x0000083D
		Public Property Field As String
			Get
				Return Me._field
			End Get
			Set(value As String)
				Me._field = value
			End Set
		End Property

		' Token: 0x1700002D RID: 45
		' (get) Token: 0x0600007D RID: 125 RVA: 0x00002646 File Offset: 0x00000846
		' (set) Token: 0x0600007E RID: 126 RVA: 0x0000264E File Offset: 0x0000084E
		Public Property Value As Object
			Get
				Return Me._value
			End Get
			Set(value As Object)
				Me._value = value
			End Set
		End Property

		' Token: 0x0600007F RID: 127 RVA: 0x00002657 File Offset: 0x00000857
		Public Sub New(field As String, value As Object)
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
			Me.Field = field
			Me.Value = value
		End Sub

		' Token: 0x04000035 RID: 53
		Private _field As String

		' Token: 0x04000036 RID: 54
		Private _value As Object
	End Class
End Namespace
