Imports System

Namespace 小条先切
	' Token: 0x02000013 RID: 19
	Public Class User
		' Token: 0x17000035 RID: 53
		' (get) Token: 0x060000C5 RID: 197 RVA: 0x000027D7 File Offset: 0x000009D7
		' (set) Token: 0x060000C6 RID: 198 RVA: 0x000027DF File Offset: 0x000009DF
		Public Property ID As Long
			Get
				Return Me._iD
			End Get
			Set(value As Long)
				Me._iD = value
			End Set
		End Property

		' Token: 0x17000036 RID: 54
		' (get) Token: 0x060000C7 RID: 199 RVA: 0x000027E8 File Offset: 0x000009E8
		' (set) Token: 0x060000C8 RID: 200 RVA: 0x000027F0 File Offset: 0x000009F0
		Public Property HardInfo As String
			Get
				Return Me._hardInfo
			End Get
			Set(value As String)
				Me._hardInfo = value
			End Set
		End Property

		' Token: 0x17000037 RID: 55
		' (get) Token: 0x060000C9 RID: 201 RVA: 0x000027F9 File Offset: 0x000009F9
		' (set) Token: 0x060000CA RID: 202 RVA: 0x00002801 File Offset: 0x00000A01
		Public Property OrderDate As Long
			Get
				Return Me._orderDate
			End Get
			Set(value As Long)
				Me._orderDate = value
			End Set
		End Property

		' Token: 0x17000038 RID: 56
		' (get) Token: 0x060000CB RID: 203 RVA: 0x0000280A File Offset: 0x00000A0A
		' (set) Token: 0x060000CC RID: 204 RVA: 0x00002812 File Offset: 0x00000A12
		Public Property DueDate As Long
			Get
				Return Me._dueDate
			End Get
			Set(value As Long)
				Me._dueDate = value
			End Set
		End Property

		' Token: 0x17000039 RID: 57
		' (get) Token: 0x060000CD RID: 205 RVA: 0x0000281B File Offset: 0x00000A1B
		' (set) Token: 0x060000CE RID: 206 RVA: 0x00002823 File Offset: 0x00000A23
		Public Property LastLoginDate As Long
			Get
				Return Me._lastLoginDate
			End Get
			Set(value As Long)
				Me._lastLoginDate = value
			End Set
		End Property

		' Token: 0x1700003A RID: 58
		' (get) Token: 0x060000CF RID: 207 RVA: 0x0000282C File Offset: 0x00000A2C
		' (set) Token: 0x060000D0 RID: 208 RVA: 0x00002834 File Offset: 0x00000A34
		Public Property AuthorGuy As String
			Get
				Return Me._authorGuy
			End Get
			Set(value As String)
				Me._authorGuy = value
			End Set
		End Property

		' Token: 0x1700003B RID: 59
		' (get) Token: 0x060000D1 RID: 209 RVA: 0x0000283D File Offset: 0x00000A3D
		' (set) Token: 0x060000D2 RID: 210 RVA: 0x00002845 File Offset: 0x00000A45
		Public Property RemianDays As Integer
			Get
				Return Me._remianDays
			End Get
			Set(value As Integer)
				Me._remianDays = value
			End Set
		End Property

		' Token: 0x1700003C RID: 60
		' (get) Token: 0x060000D3 RID: 211 RVA: 0x0000284E File Offset: 0x00000A4E
		' (set) Token: 0x060000D4 RID: 212 RVA: 0x00002856 File Offset: 0x00000A56
		Public Property IsLoginTody As Boolean
			Get
				Return Me._IsLoginTody
			End Get
			Set(value As Boolean)
				Me._IsLoginTody = value
			End Set
		End Property

		' Token: 0x060000D5 RID: 213 RVA: 0x00002395 File Offset: 0x00000595
		Public Sub New()
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
		End Sub

		' Token: 0x04000061 RID: 97
		Private _iD As Long

		' Token: 0x04000062 RID: 98
		Private _hardInfo As String

		' Token: 0x04000063 RID: 99
		Private _orderDate As Long

		' Token: 0x04000064 RID: 100
		Private _dueDate As Long

		' Token: 0x04000065 RID: 101
		Private _lastLoginDate As Long

		' Token: 0x04000066 RID: 102
		Private _authorGuy As String

		' Token: 0x04000067 RID: 103
		Private _remianDays As Integer

		' Token: 0x04000068 RID: 104
		Private _IsLoginTody As Boolean
	End Class
End Namespace
