Imports System
Imports AlphaCAMRouter

Namespace 小条先切
	' Token: 0x02000004 RID: 4
	Public Class UniqueGeo
		' Token: 0x17000007 RID: 7
		' (get) Token: 0x06000013 RID: 19 RVA: 0x0000232F File Offset: 0x0000052F
		' (set) Token: 0x06000014 RID: 20 RVA: 0x00002337 File Offset: 0x00000537
		Public Property Path As Path
			Get
				Return Me._path
			End Get
			Set(value As Path)
				Me._path = value
			End Set
		End Property

		' Token: 0x17000008 RID: 8
		' (get) Token: 0x06000015 RID: 21 RVA: 0x00002340 File Offset: 0x00000540
		' (set) Token: 0x06000016 RID: 22 RVA: 0x00002348 File Offset: 0x00000548
		Public Property IsDelete As Boolean
			Get
				Return Me._isDelete
			End Get
			Set(value As Boolean)
				Me._isDelete = value
			End Set
		End Property

		' Token: 0x17000009 RID: 9
		' (get) Token: 0x06000017 RID: 23 RVA: 0x00002351 File Offset: 0x00000551
		' (set) Token: 0x06000018 RID: 24 RVA: 0x00002359 File Offset: 0x00000559
		Public Property StartX As Double
			Get
				Return Me._startX
			End Get
			Set(value As Double)
				Me._startX = value
			End Set
		End Property

		' Token: 0x1700000A RID: 10
		' (get) Token: 0x06000019 RID: 25 RVA: 0x00002362 File Offset: 0x00000562
		' (set) Token: 0x0600001A RID: 26 RVA: 0x0000236A File Offset: 0x0000056A
		Public Property StartY As Double
			Get
				Return Me._startY
			End Get
			Set(value As Double)
				Me._startY = value
			End Set
		End Property

		' Token: 0x1700000B RID: 11
		' (get) Token: 0x0600001B RID: 27 RVA: 0x00002373 File Offset: 0x00000573
		' (set) Token: 0x0600001C RID: 28 RVA: 0x0000237B File Offset: 0x0000057B
		Public Property EndX As Double
			Get
				Return Me._endX
			End Get
			Set(value As Double)
				Me._endX = value
			End Set
		End Property

		' Token: 0x1700000C RID: 12
		' (get) Token: 0x0600001D RID: 29 RVA: 0x00002384 File Offset: 0x00000584
		' (set) Token: 0x0600001E RID: 30 RVA: 0x0000238C File Offset: 0x0000058C
		Public Property EndY As Double
			Get
				Return Me._endY
			End Get
			Set(value As Double)
				Me._endY = value
			End Set
		End Property

		' Token: 0x0600001F RID: 31 RVA: 0x00002B28 File Offset: 0x00000D28
		Public Function UniquePath(TGTest As UniqueGeo, ByRef SX As Double, ByRef SY As Double, ByRef EX As Double, ByRef EY As Double) As Boolean
			Dim startX As Double = Me.StartX
			Dim startY As Double = Me.StartY
			Dim endX As Double = Me.EndX
			Dim endY As Double = Me.EndY
			Dim startX2 As Double = TGTest.StartX
			Dim startY2 As Double = TGTest.StartY
			Dim endX2 As Double = TGTest.EndX
			Dim endY2 As Double = TGTest.EndY
			Dim flag As Boolean
			If startX <> endX AndAlso startY <= endY + 0.0001 AndAlso startY >= endY - 0.0001 Then
				flag = True
			Else
				If startX > endX + 0.0001 OrElse startX < endX - 0.0001 OrElse startY = endY Then
					Return False
				End If
				flag = False
			End If
			Dim flag2 As Boolean
			If startX2 <> endX2 AndAlso startY2 <= endY2 + 0.0001 AndAlso startY2 >= endY2 - 0.0001 Then
				flag2 = True
			Else
				If startX2 > endX2 + 0.0001 OrElse startX2 < endX2 - 0.0001 OrElse startY2 = endY2 Then
					Return False
				End If
				flag2 = False
			End If
			If flag <> flag2 Then
				Return False
			End If
			If flag Then
				If startY > startY2 + 0.0001 OrElse startY < startY2 - 0.0001 Then
					Return False
				End If
				If startX2 > endX + 0.0001 + 15.0 OrElse endX2 < startX - 0.0001 - 15.0 Then
					Return False
				End If
				SY = endY
				EY = endY
				SX = If((startX < startX2), startX, startX2)
				EX = If((endX > endX2), endX, endX2)
			Else
				If startX > startX2 + 0.0001 OrElse startX < startX2 - 0.0001 Then
					Return False
				End If
				If startY2 > endY + 0.0001 + 15.0 OrElse endY2 < startY - 0.0001 - 15.0 Then
					Return False
				End If
				SX = startX
				EX = startX
				SY = If((startY < startY2), startY, startY2)
				EY = If((endY > endY2), endY, endY2)
			End If
			Return True
		End Function

		' Token: 0x06000020 RID: 32 RVA: 0x00002395 File Offset: 0x00000595
		Public Sub New()
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
		End Sub

		' Token: 0x04000008 RID: 8
		Private _path As Path

		' Token: 0x04000009 RID: 9
		Private _isDelete As Boolean

		' Token: 0x0400000A RID: 10
		Private _startX As Double

		' Token: 0x0400000B RID: 11
		Private _startY As Double

		' Token: 0x0400000C RID: 12
		Private _endX As Double

		' Token: 0x0400000D RID: 13
		Private _endY As Double
	End Class
End Namespace
