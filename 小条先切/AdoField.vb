Imports System

Namespace 小条先切
	' Token: 0x0200000A RID: 10
	Public Class AdoField
		' Token: 0x1700002E RID: 46
		' (get) Token: 0x06000080 RID: 128 RVA: 0x00002672 File Offset: 0x00000872
		' (set) Token: 0x06000081 RID: 129 RVA: 0x0000267A File Offset: 0x0000087A
		Public Property Field As String
			Get
				Return Me._field
			End Get
			Set(value As String)
				Me._field = value
			End Set
		End Property

		' Token: 0x1700002F RID: 47
		' (get) Token: 0x06000082 RID: 130 RVA: 0x00002683 File Offset: 0x00000883
		' (set) Token: 0x06000083 RID: 131 RVA: 0x0000268B File Offset: 0x0000088B
		Public Property Value As Object
			Get
				Return Me._value
			End Get
			Set(value As Object)
				Me._value = value
			End Set
		End Property

		' Token: 0x06000084 RID: 132 RVA: 0x00002694 File Offset: 0x00000894
		Public Sub New(field As String, value As Object)
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
			Me.Field = field
			Me.Value = value
		End Sub

		' Token: 0x04000037 RID: 55
		Private _field As String

		' Token: 0x04000038 RID: 56
		Private _value As Object
	End Class
End Namespace
