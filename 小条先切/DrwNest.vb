Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports AlphaCAMRouter

Namespace 小条先切
	' Token: 0x02000006 RID: 6
	Public Class DrwNest
		' Token: 0x1700001D RID: 29
		' (get) Token: 0x06000046 RID: 70 RVA: 0x000024B9 File Offset: 0x000006B9
		' (set) Token: 0x06000047 RID: 71 RVA: 0x000024C1 File Offset: 0x000006C1
		Public Property NestName As String
			Get
				Return Me._nestName
			End Get
			Set(value As String)
				Me._nestName = value
			End Set
		End Property

		' Token: 0x1700001E RID: 30
		' (get) Token: 0x06000048 RID: 72 RVA: 0x000024CA File Offset: 0x000006CA
		' (set) Token: 0x06000049 RID: 73 RVA: 0x000024D2 File Offset: 0x000006D2
		Public Property SheetNo As Integer
			Get
				Return Me._sheetNo
			End Get
			Set(value As Integer)
				Me._sheetNo = value
			End Set
		End Property

		' Token: 0x1700001F RID: 31
		' (get) Token: 0x0600004A RID: 74 RVA: 0x000024DB File Offset: 0x000006DB
		' (set) Token: 0x0600004B RID: 75 RVA: 0x000024E3 File Offset: 0x000006E3
		Public Property Thickness As Double
			Get
				Return Me._thickness
			End Get
			Set(value As Double)
				Me._thickness = value
			End Set
		End Property

		' Token: 0x17000020 RID: 32
		' (get) Token: 0x0600004C RID: 76 RVA: 0x000024EC File Offset: 0x000006EC
		' (set) Token: 0x0600004D RID: 77 RVA: 0x000024F4 File Offset: 0x000006F4
		Public Property MinX As Double
			Get
				Return Me._minX
			End Get
			Set(value As Double)
				Me._minX = value
			End Set
		End Property

		' Token: 0x17000021 RID: 33
		' (get) Token: 0x0600004E RID: 78 RVA: 0x000024FD File Offset: 0x000006FD
		' (set) Token: 0x0600004F RID: 79 RVA: 0x00002505 File Offset: 0x00000705
		Public Property MinY As Double
			Get
				Return Me._minY
			End Get
			Set(value As Double)
				Me._minY = value
			End Set
		End Property

		' Token: 0x17000022 RID: 34
		' (get) Token: 0x06000050 RID: 80 RVA: 0x0000250E File Offset: 0x0000070E
		' (set) Token: 0x06000051 RID: 81 RVA: 0x00002516 File Offset: 0x00000716
		Public Property MaxX As Double
			Get
				Return Me._maxX
			End Get
			Set(value As Double)
				Me._maxX = value
			End Set
		End Property

		' Token: 0x17000023 RID: 35
		' (get) Token: 0x06000052 RID: 82 RVA: 0x0000251F File Offset: 0x0000071F
		' (set) Token: 0x06000053 RID: 83 RVA: 0x00002527 File Offset: 0x00000727
		Public Property MaxY As Double
			Get
				Return Me._maxY
			End Get
			Set(value As Double)
				Me._maxY = value
			End Set
		End Property

		' Token: 0x17000024 RID: 36
		' (get) Token: 0x06000054 RID: 84 RVA: 0x00002530 File Offset: 0x00000730
		' (set) Token: 0x06000055 RID: 85 RVA: 0x00002538 File Offset: 0x00000738
		Public Property ToolPaths As List(Of CutToolPath)
			Get
				Return Me._toolPaths
			End Get
			Set(value As List(Of CutToolPath))
				Me._toolPaths = value
			End Set
		End Property

		' Token: 0x06000056 RID: 86 RVA: 0x00002541 File Offset: 0x00000741
		Public Sub New()
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
			Me.ToolPaths = New List(Of CutToolPath)()
		End Sub

		' Token: 0x06000057 RID: 87 RVA: 0x00002559 File Offset: 0x00000759
		Public Function TextInside(X As Double, Y As Double) As Boolean
			Return X >= Me.MinX AndAlso X <= Me.MaxX AndAlso Y >= Me.MinY AndAlso Y <= Me.MaxY
		End Function

		' Token: 0x17000025 RID: 37
		' (get) Token: 0x06000058 RID: 88 RVA: 0x00002582 File Offset: 0x00000782
		' (set) Token: 0x06000059 RID: 89 RVA: 0x0000258A File Offset: 0x0000078A
		Public Property Acam As App
			Get
				Return Me._acam
			End Get
			Set(value As App)
				Me._acam = value
			End Set
		End Property

		' Token: 0x17000026 RID: 38
		' (get) Token: 0x0600005A RID: 90 RVA: 0x00002593 File Offset: 0x00000793
		' (set) Token: 0x0600005B RID: 91 RVA: 0x0000259B File Offset: 0x0000079B
		Public Property SloopingLine As Boolean
			Get
				Return Me._sloopingLine
			End Get
			Set(value As Boolean)
				Me._sloopingLine = value
			End Set
		End Property

		' Token: 0x17000027 RID: 39
		' (get) Token: 0x0600005C RID: 92 RVA: 0x000025A4 File Offset: 0x000007A4
		' (set) Token: 0x0600005D RID: 93 RVA: 0x000025AC File Offset: 0x000007AC
		Public Property IsNestH As Boolean
			Get
				Return Me._isNestH
			End Get
			Set(value As Boolean)
				Me._isNestH = value
			End Set
		End Property

		' Token: 0x17000028 RID: 40
		' (get) Token: 0x0600005E RID: 94 RVA: 0x000025B5 File Offset: 0x000007B5
		' (set) Token: 0x0600005F RID: 95 RVA: 0x000025BD File Offset: 0x000007BD
		Public Property BackDist As Double
			Get
				Return Me._backDist
			End Get
			Set(value As Double)
				Me._backDist = value
			End Set
		End Property

		' Token: 0x17000029 RID: 41
		' (get) Token: 0x06000060 RID: 96 RVA: 0x000025C6 File Offset: 0x000007C6
		' (set) Token: 0x06000061 RID: 97 RVA: 0x000025CE File Offset: 0x000007CE
		Public Property SmallPanel As Double
			Get
				Return Me._smallPanel
			End Get
			Set(value As Double)
				Me._smallPanel = value
			End Set
		End Property

		' Token: 0x1700002A RID: 42
		' (get) Token: 0x06000062 RID: 98 RVA: 0x000025D7 File Offset: 0x000007D7
		' (set) Token: 0x06000063 RID: 99 RVA: 0x000025DF File Offset: 0x000007DF
		Public Property SloopDist As Double
			Get
				Return Me._sloopDist
			End Get
			Set(value As Double)
				Me._sloopDist = value
			End Set
		End Property

		' Token: 0x1700002B RID: 43
		' (get) Token: 0x06000064 RID: 100 RVA: 0x000025E8 File Offset: 0x000007E8
		' (set) Token: 0x06000065 RID: 101 RVA: 0x000025F0 File Offset: 0x000007F0
		Public Property UniqueCutTool As Boolean
			Get
				Return Me._uniqueCutTool
			End Get
			Set(value As Boolean)
				Me._uniqueCutTool = value
			End Set
		End Property

		' Token: 0x06000066 RID: 102 RVA: 0x00004088 File Offset: 0x00002288
		Public Sub OptimizeCutPaths()
			Dim toolPaths As List(Of CutToolPath) = Me.ToolPaths
			If toolPaths.Count <= 0 Then
				Return
			End If
			Me.Acam.SelectTool(toolPaths(0).ToolFileName)
			For i As Integer = 0 To toolPaths.Count - 1
				Dim num As Double = toolPaths.ElementAt(i).MaxX - toolPaths.ElementAt(i).MinX
				Dim num2 As Double = toolPaths.ElementAt(i).MaxY - toolPaths.ElementAt(i).MinY
				If(toolPaths.ElementAt(i).Horizontal AndAlso num2 <= Me.SmallPanel) OrElse (Not toolPaths.ElementAt(i).Horizontal AndAlso num <= Me.SmallPanel) Then
					toolPaths.ElementAt(i).DepthCutSpecial = 1
				End If
				toolPaths.ElementAt(i).DrawToolGeo(Me.Acam.ActiveDrawing)
			Next
			Dim list As List(Of CutToolPath) = New List(Of CutToolPath)()
			Dim list2 As List(Of CutToolPath) = New List(Of CutToolPath)()
			Dim count As Integer = toolPaths.Count
			For j As Integer = 0 To count - 1
				Me.SetCutPriority(toolPaths)
				Me.SortToolPathByAxis(toolPaths, Me.IsNestH)
				Me.SortToolPathByCutSpecial(toolPaths)
				Me.SortToolPathByPriority(toolPaths)
				If toolPaths.ElementAt(0).DepthCutSpecial >= 1 Then
					list.Add(toolPaths.ElementAt(0))
				Else
					list2.Add(toolPaths.ElementAt(0))
				End If
				toolPaths.RemoveAt(0)
			Next
			For k As Integer = 0 To list.Count - 1
				list(k).RoughFinish(Me.Acam, Me.SloopingLine, Me.BackDist, Me.SloopDist, 0.5)
				list(k).ToolGeo.Delete()
				list(k).ToolPath.Delete()
			Next
			Dim minX As Double = Me.MinX
			Dim minY As Double = Me.MinY
			If Me.UniqueCutTool Then
				Me.UniqueToolPath(list2)
				For l As Integer = 0 To list2.Count - 1
					list2(l).ToolGeo.Delete()
					list2(l).ToolPath.Delete()
				Next
				Return
			End If
			For m As Integer = 0 To list2.Count - 1
				list2(m).RoughFinishBigPanel(Me.Acam, minX, minY, minX, minY)
				list2(m).ToolGeo.Delete()
				list2(m).ToolPath.Delete()
			Next
		End Sub

		' Token: 0x06000067 RID: 103 RVA: 0x000042F8 File Offset: 0x000024F8
		Private Sub UniqueToolPath(lstToolPaht As List(Of CutToolPath))
			If lstToolPaht.Count <= 0 Then
				Return
			End If
			Dim millData As MillData = lstToolPaht(0).MillData
			Dim activeDrawing As Drawing = Me.Acam.ActiveDrawing
			Dim list As List(Of UniqueGeo) = New List(Of UniqueGeo)()
			Dim list2 As List(Of UniqueGeo) = New List(Of UniqueGeo)()
			Dim list3 As List(Of UniqueGeo) = New List(Of UniqueGeo)()
			For i As Integer = 0 To lstToolPaht.Count - 1
				Dim elements As Elements = lstToolPaht(i).ToolGeo.Elements
				Dim count As Integer = elements.Count
				For j As Integer = 1 To count
					Dim element As Element = elements.Item(j)
					Dim uniqueGeo As UniqueGeo = New UniqueGeo()
					If element.IsLine Then
						Dim num As Double = element.StartXL
						Dim num2 As Double = element.StartYL
						Dim num3 As Double = element.EndXL
						Dim num4 As Double = element.EndYL
						If num > num3 OrElse num2 > num4 Then
							Dim num5 As Double = num
							Dim num6 As Double = num2
							num = num3
							num2 = num4
							num3 = num5
							num4 = num6
						End If
						If num <= num3 + 0.0001 AndAlso num >= num3 - 0.0001 Then
							uniqueGeo.StartX = num
							uniqueGeo.StartY = num2
							uniqueGeo.EndX = num3
							uniqueGeo.EndY = num4
							list.Add(uniqueGeo)
						ElseIf num2 <= num4 + 0.0001 AndAlso num2 >= num4 - 0.0001 Then
							uniqueGeo.StartX = num
							uniqueGeo.StartY = num2
							uniqueGeo.EndX = num3
							uniqueGeo.EndY = num4
							list2.Add(uniqueGeo)
						Else
							uniqueGeo.Path = activeDrawing.Create2DLine(num, num2, num3, num4)
							uniqueGeo.Path.ToolInOut = AcamToolInOut.acamON_CENTER
							uniqueGeo.Path.ToolSide = AcamToolSide.const_1
							list3.Add(uniqueGeo)
						End If
					Else
						Dim geo2D As Geo2D = activeDrawing.Create2DGeometry(element.StartXL, element.StartYL)
						geo2D.AddArcPointCenter(element.EndXL, element.EndYL, element.CenterXL, element.CenterYL, element.CW)
						uniqueGeo.Path = geo2D.Finish()
						uniqueGeo.Path.ToolInOut = AcamToolInOut.acamON_CENTER
						uniqueGeo.Path.ToolSide = AcamToolSide.const_1
						list3.Add(uniqueGeo)
					End If
				Next
			Next
			Me.UniqueLine(list)
			Me.UniqueLine(list2)
			Dim list4 As List(Of UniqueGeo) = New List(Of UniqueGeo)()
			For k As Integer = 0 To list.Count - 1
				If Not list(k).IsDelete Then
					list4.Add(New UniqueGeo() With { .IsDelete = False, .Path = activeDrawing.Create2DLine(list(k).StartX, list(k).StartY, list(k).EndX, list(k).EndY) })
				End If
			Next
			For l As Integer = 0 To list2.Count - 1
				If Not list2(l).IsDelete Then
					list4.Add(New UniqueGeo() With { .IsDelete = False, .Path = activeDrawing.Create2DLine(list2(l).StartX, list2(l).StartY, list2(l).EndX, list2(l).EndY) })
				End If
			Next
			For m As Integer = 0 To list4.Count - 1
				list4(m).Path.Selected = True
			Next
			For n As Integer = 0 To list3.Count - 1
				list3(n).Path.Selected = True
			Next
			Dim paths As Paths = activeDrawing.Join()
			list4.Clear()
			For num7 As Integer = 1 To paths.Count
				list4.Add(New UniqueGeo() With { .Path = paths.Item(num7) })
			Next
			Me.SortLineByY(list4)
			Me.SortLineByX(list4)
			Me.SortLineByLength(list4)
			Dim finalDepth As Single = millData.FinalDepth
			For num8 As Integer = 0 To list4.Count - 1
				If num8 > 0 AndAlso Me.GetReverse(list4(num8 - 1), list4(num8)) Then
					list4(num8).Path.Reverse()
				End If
				If finalDepth <= -20F Then
					list4(num8).Path.Selected = True
					millData.FinalDepth = CSng((CDbl(finalDepth) * 0.6))
					millData.RoughFinish()
					list4(num8).Path.Reverse()
				End If
				list4(num8).Path.Selected = True
				millData.FinalDepth = finalDepth
				millData.RoughFinish()
			Next
			For num9 As Integer = 0 To list4.Count - 1
				list4(num9).Path.Delete()
			Next
		End Sub

		' Token: 0x06000068 RID: 104 RVA: 0x000047EC File Offset: 0x000029EC
		Private Sub UniqueLine(lines As List(Of UniqueGeo))
			Dim num As Double = 0.0
			Dim num2 As Double = CDbl(0F)
			Dim num3 As Double = CDbl(0F)
			Dim num4 As Double = CDbl(0F)
			Dim num5 As Double = num
			For i As Integer = 0 To lines.Count - 1
				If Not lines(i).IsDelete Then
					For j As Integer = 0 To lines.Count - 1
						If Not lines(j).IsDelete AndAlso i <> j AndAlso lines(i).UniquePath(lines(j), num5, num3, num4, num2) Then
							If i <= j Then
								lines(j).StartX = num5
								lines(j).StartY = num3
								lines(j).EndX = num4
								lines(j).EndY = num2
								lines(i).IsDelete = True
								Exit For
							End If
							lines(i).StartX = num5
							lines(i).StartY = num3
							lines(i).EndX = num4
							lines(i).EndY = num2
							lines(j).IsDelete = True
						End If
					Next
				End If
			Next
		End Sub

		' Token: 0x06000069 RID: 105 RVA: 0x00004928 File Offset: 0x00002B28
		Private Sub SortLineByLength(lines As List(Of UniqueGeo))
			For i As Integer = 0 To lines.Count - 1 - 1
				For j As Integer = 0 To lines.Count - 1 - i - 1
					If lines(j).Path.Length >= 2.0 * Me.SmallPanel AndAlso lines(j).Path.Length > lines(j + 1).Path.Length Then
						Dim path As Path = lines(j + 1).Path
						lines(j + 1).Path = lines(j).Path
						lines(j).Path = path
					End If
				Next
			Next
		End Sub

		' Token: 0x0600006A RID: 106 RVA: 0x000049E8 File Offset: 0x00002BE8
		Private Sub SortLineByX(lines As List(Of UniqueGeo))
			For i As Integer = 0 To lines.Count - 1 - 1
				For j As Integer = 0 To lines.Count - 1 - i - 1
					Dim minXL As Double = lines(j).Path.MinXL
					Dim minXL2 As Double = lines(j + 1).Path.MinXL
					If minXL > minXL2 + Me.SmallPanel Then
						Dim path As Path = lines(j + 1).Path
						lines(j + 1).Path = lines(j).Path
						lines(j).Path = path
					End If
				Next
			Next
		End Sub

		' Token: 0x0600006B RID: 107 RVA: 0x00004A8C File Offset: 0x00002C8C
		Private Sub SortLineByY(lines As List(Of UniqueGeo))
			For i As Integer = 0 To lines.Count - 1 - 1
				For j As Integer = 0 To lines.Count - 1 - i - 1
					Dim minYL As Double = lines(j).Path.MinYL
					Dim minYL2 As Double = lines(j + 1).Path.MinYL
					If minYL > minYL2 + Me.SmallPanel Then
						Dim path As Path = lines(j + 1).Path
						lines(j + 1).Path = lines(j).Path
						lines(j).Path = path
					End If
				Next
			Next
		End Sub

		' Token: 0x0600006C RID: 108 RVA: 0x00004B30 File Offset: 0x00002D30
		Private Function GetReverse(p1 As UniqueGeo, p2 As UniqueGeo) As Boolean
			Dim lastElem As Element = p1.Path.GetLastElem()
			Dim endXL As Double = lastElem.EndXL
			Dim endYL As Double = lastElem.EndYL
			Dim firstElem As Element = p2.Path.GetFirstElem()
			Dim startXL As Double = firstElem.StartXL
			Dim startYL As Double = firstElem.StartYL
			Dim lastElem2 As Element = p2.Path.GetLastElem()
			Dim endXL2 As Double = lastElem2.EndXL
			Dim endYL2 As Double = lastElem2.EndYL
			Return(endXL - startXL) * (endXL - startXL) + (endYL - startYL) * (endYL - startYL) > (endXL - endXL2) * (endXL - endXL2) + (endYL - endYL2) * (endYL - endYL2)
		End Function

		' Token: 0x0600006D RID: 109 RVA: 0x00004BC0 File Offset: 0x00002DC0
		Private Sub SortToolPathByCutSpecial(tps As List(Of CutToolPath))
			For i As Integer = 0 To tps.Count - 1 - 1
				For j As Integer = 0 To tps.Count - i - 1 - 1
					If tps.ElementAt(j).DepthCutSpecial < tps.ElementAt(j + 1).DepthCutSpecial Then
						Dim cutToolPath As CutToolPath = tps.ElementAt(j)
						tps(j) = tps.ElementAt(j + 1)
						tps(j + 1) = cutToolPath
					End If
				Next
			Next
		End Sub

		' Token: 0x0600006E RID: 110 RVA: 0x00004C34 File Offset: 0x00002E34
		Private Sub SortToolPathByPriority(tps As List(Of CutToolPath))
			For i As Integer = 0 To tps.Count - 1 - 1
				For j As Integer = 0 To tps.Count - i - 1 - 1
					If tps.ElementAt(j).CutPriority < tps.ElementAt(j + 1).CutPriority Then
						Dim cutToolPath As CutToolPath = tps.ElementAt(j)
						tps(j) = tps.ElementAt(j + 1)
						tps(j + 1) = cutToolPath
					End If
				Next
			Next
		End Sub

		' Token: 0x0600006F RID: 111 RVA: 0x00004CA8 File Offset: 0x00002EA8
		Private Sub SortToolPathByAxis(tps As List(Of CutToolPath), isNestH As Boolean)
			For i As Integer = 0 To tps.Count - 1 - 1
				For j As Integer = 0 To tps.Count - i - 1 - 1
					If isNestH Then
						If tps.ElementAt(j).MaxX > tps.ElementAt(j + 1).MaxX Then
							Dim cutToolPath As CutToolPath = tps.ElementAt(j)
							tps(j) = tps.ElementAt(j + 1)
							tps(j + 1) = cutToolPath
						End If
					ElseIf tps.ElementAt(j).MaxY > tps.ElementAt(j + 1).MaxY Then
						Dim cutToolPath2 As CutToolPath = tps.ElementAt(j)
						tps(j) = tps.ElementAt(j + 1)
						tps(j + 1) = cutToolPath2
					End If
				Next
			Next
		End Sub

		' Token: 0x06000070 RID: 112 RVA: 0x00004D6C File Offset: 0x00002F6C
		Private Sub SetCutPriority(tps As List(Of CutToolPath))
			For i As Integer = 0 To tps.Count - 1
				tps.ElementAt(i).CutPriority = 0
				tps.ElementAt(i).CutDirect = 0
			Next
			Dim j As Integer = 0
			IL_06E7:
			While j < tps.Count
				Dim num As Integer = 0
				Dim num2 As Integer = 0
				Dim num3 As Double = tps.ElementAt(j).MinX
				Dim num4 As Double = tps.ElementAt(j).MinY
				Dim num5 As Double = tps.ElementAt(j).MaxX
				Dim num6 As Double = tps.ElementAt(j).MaxY
				num3 += (num5 - num3) * 0.1
				num5 -= (num5 - num3) * 0.1
				num4 += (num6 - num4) * 0.1
				num6 -= (num6 - num4) * 0.1
				Dim flag As Boolean = False
				For k As Integer = 0 To tps.Count - 1
					If j <> k Then
						Dim minX As Double = tps.ElementAt(k).MinX
						Dim minY As Double = tps.ElementAt(k).MinY
						Dim maxX As Double = tps.ElementAt(k).MaxX
						Dim maxY As Double = tps.ElementAt(k).MaxY
						If minY < num6 AndAlso maxY > num4 AndAlso minX < num5 Then
							If maxX > num3 Then
								Dim path As Path = tps.ElementAt(j).ToolGeo.Copy()
								path.Visible = False
								path.ScaleL2((num5 - num3) / (path.MaxXL - path.MinXL), (num6 - num4) / (path.MaxYL - path.MinYL), (path.MaxXL + path.MinXL) / 2.0, (path.MaxYL + path.MinYL) / 2.0)
								path.MoveL(-(num5 - num3) / 2.0, 0.0)
								If Not path.TestIntersectPath(tps.ElementAt(k).ToolGeo, 0.0, 0.0) Then
									path.Delete()
									GoTo IL_01FE
								End If
								flag = True
								path.Delete()
							Else
								flag = True
							End If
							IL_0222:
							If Not flag Then
								num2 += 1
								num += If(tps.ElementAt(j).Horizontal, 1, 3)
							End If
							Dim flag2 As Boolean = False
							For l As Integer = 0 To tps.Count - 1
								If j <> l Then
									Dim minX2 As Double = tps.ElementAt(l).MinX
									Dim minY2 As Double = tps.ElementAt(l).MinY
									Dim maxX2 As Double = tps.ElementAt(l).MaxX
									Dim maxY2 As Double = tps.ElementAt(l).MaxY
									If maxY2 > num4 AndAlso minX2 < num5 AndAlso maxX2 > num3 Then
										If minY2 < num6 Then
											Dim path2 As Path = tps.ElementAt(j).ToolGeo.Copy()
											path2.Visible = False
											path2.ScaleL2((num5 - num3) / (path2.MaxXL - path2.MinXL), (num6 - num4) / (path2.MaxYL - path2.MinYL), (path2.MaxXL + path2.MinXL) / 2.0, (path2.MaxYL + path2.MinYL) / 2.0)
											path2.MoveL(0.0, (num6 - num4) / 2.0)
											If Not path2.TestIntersectPath(tps.ElementAt(l).ToolGeo, 0.0, 0.0) Then
												path2.Delete()
												GoTo IL_037E
											End If
											flag2 = True
											path2.Delete()
										Else
											flag2 = True
										End If
										IL_03A2:
										If Not flag2 Then
											num2 += 2
											num += If(tps.ElementAt(j).Horizontal, 3, 1)
										End If
										Dim flag3 As Boolean = False
										For m As Integer = 0 To tps.Count - 1
											If j <> m Then
												Dim minX3 As Double = tps.ElementAt(m).MinX
												Dim minY3 As Double = tps.ElementAt(m).MinY
												Dim maxX3 As Double = tps.ElementAt(m).MaxX
												Dim maxY3 As Double = tps.ElementAt(m).MaxY
												If minY3 < num6 AndAlso maxY3 > num4 AndAlso maxX3 > num3 Then
													If minX3 < num5 Then
														Dim path3 As Path = tps.ElementAt(j).ToolGeo.Copy()
														path3.Visible = False
														path3.ScaleL2((num5 - num3) / (path3.MaxXL - path3.MinXL), (num6 - num4) / (path3.MaxYL - path3.MinYL), (path3.MaxXL + path3.MinXL) / 2.0, (path3.MaxYL + path3.MinYL) / 2.0)
														path3.MoveL((num5 - num3) / 2.0, 0.0)
														If Not path3.TestIntersectPath(tps.ElementAt(m).ToolGeo, 0.0, 0.0) Then
															path3.Delete()
															GoTo IL_04FE
														End If
														flag3 = True
														path3.Delete()
													Else
														flag3 = True
													End If
													IL_0522:
													If Not flag3 Then
														num2 += 4
														num += If(tps.ElementAt(j).Horizontal, 1, 3)
													End If
													Dim flag4 As Boolean = False
													For n As Integer = 0 To tps.Count - 1
														If j <> n Then
															Dim minX4 As Double = tps.ElementAt(n).MinX
															Dim minY4 As Double = tps.ElementAt(n).MinY
															Dim maxX4 As Double = tps.ElementAt(n).MaxX
															Dim maxY4 As Double = tps.ElementAt(n).MaxY
															If minY4 < num6 AndAlso minX4 < num5 AndAlso maxX4 > num3 AndAlso maxX4 > num3 Then
																If maxY4 > num4 Then
																	Dim path4 As Path = tps.ElementAt(j).ToolGeo.Copy()
																	path4.Visible = False
																	path4.ScaleL2((num5 - num3) / (path4.MaxXL - path4.MinXL), (num6 - num4) / (path4.MaxYL - path4.MinYL), (path4.MaxXL + path4.MinXL) / 2.0, (path4.MaxYL + path4.MinYL) / 2.0)
																	path4.MoveL(0.0, -(num6 - num4) / 2.0)
																	If Not path4.TestIntersectPath(tps.ElementAt(n).ToolGeo, 0.0, 0.0) Then
																		path4.Delete()
																		GoTo IL_0688
																	End If
																	flag4 = True
																	path4.Delete()
																Else
																	flag4 = True
																End If
																IL_06AC:
																If Not flag4 Then
																	num2 += 8
																	num += If(tps.ElementAt(j).Horizontal, 3, 1)
																End If
																tps.ElementAt(j).CutPriority = num
																tps.ElementAt(j).CutDirect = num2
																j += 1
																GoTo IL_06E7
															End If
														End If
														IL_0688:
													Next
													GoTo IL_06AC
												End If
											End If
											IL_04FE:
										Next
										GoTo IL_0522
									End If
								End If
								IL_037E:
							Next
							GoTo IL_03A2
						End If
					End If
					IL_01FE:
				Next
				GoTo IL_0222
			End While
		End Sub

		' Token: 0x0400001E RID: 30
		Private _nestName As String

		' Token: 0x0400001F RID: 31
		Private _sheetNo As Integer

		' Token: 0x04000020 RID: 32
		Private _thickness As Double

		' Token: 0x04000021 RID: 33
		Private _minX As Double

		' Token: 0x04000022 RID: 34
		Private _minY As Double

		' Token: 0x04000023 RID: 35
		Private _maxX As Double

		' Token: 0x04000024 RID: 36
		Private _maxY As Double

		' Token: 0x04000025 RID: 37
		Private _toolPaths As List(Of CutToolPath)

		' Token: 0x04000026 RID: 38
		Private _acam As App

		' Token: 0x04000027 RID: 39
		Private _sloopingLine As Boolean

		' Token: 0x04000028 RID: 40
		Private _isNestH As Boolean

		' Token: 0x04000029 RID: 41
		Private _smallPanel As Double

		' Token: 0x0400002A RID: 42
		Private _backDist As Double

		' Token: 0x0400002B RID: 43
		Private _sloopDist As Double

		' Token: 0x0400002C RID: 44
		Private _uniqueCutTool As Boolean
	End Class
End Namespace
