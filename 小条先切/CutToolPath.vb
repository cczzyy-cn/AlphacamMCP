Imports System
Imports System.Collections.Generic
Imports AlphaCAMRouter

Namespace 小条先切
	' Token: 0x02000005 RID: 5
	Public Class CutToolPath
		' Token: 0x1700000D RID: 13
		' (get) Token: 0x06000021 RID: 33 RVA: 0x000023A2 File Offset: 0x000005A2
		' (set) Token: 0x06000022 RID: 34 RVA: 0x000023AA File Offset: 0x000005AA
		Public Property MinX As Double
			Get
				Return Me._minX
			End Get
			Set(value As Double)
				Me._minX = value
			End Set
		End Property

		' Token: 0x1700000E RID: 14
		' (get) Token: 0x06000023 RID: 35 RVA: 0x000023B3 File Offset: 0x000005B3
		' (set) Token: 0x06000024 RID: 36 RVA: 0x000023BB File Offset: 0x000005BB
		Public Property MaxX As Double
			Get
				Return Me._maxX
			End Get
			Set(value As Double)
				Me._maxX = value
			End Set
		End Property

		' Token: 0x1700000F RID: 15
		' (get) Token: 0x06000025 RID: 37 RVA: 0x000023C4 File Offset: 0x000005C4
		' (set) Token: 0x06000026 RID: 38 RVA: 0x000023CC File Offset: 0x000005CC
		Public Property MinY As Double
			Get
				Return Me._minY
			End Get
			Set(value As Double)
				Me._minY = value
			End Set
		End Property

		' Token: 0x17000010 RID: 16
		' (get) Token: 0x06000027 RID: 39 RVA: 0x000023D5 File Offset: 0x000005D5
		' (set) Token: 0x06000028 RID: 40 RVA: 0x000023DD File Offset: 0x000005DD
		Public Property MaxY As Double
			Get
				Return Me._maxY
			End Get
			Set(value As Double)
				Me._maxY = value
			End Set
		End Property

		' Token: 0x17000011 RID: 17
		' (get) Token: 0x06000029 RID: 41 RVA: 0x000023E6 File Offset: 0x000005E6
		' (set) Token: 0x0600002A RID: 42 RVA: 0x000023EE File Offset: 0x000005EE
		Public Property ToolPath As Path
			Get
				Return Me.toolPath
			End Get
			Set(value As Path)
				Me.toolPath = value
			End Set
		End Property

		' Token: 0x17000012 RID: 18
		' (get) Token: 0x0600002B RID: 43 RVA: 0x000023F7 File Offset: 0x000005F7
		' (set) Token: 0x0600002C RID: 44 RVA: 0x000023FF File Offset: 0x000005FF
		Public Property InnerToolPaths As List(Of CutToolPath)
			Get
				Return Me._innerToolPaths
			End Get
			Set(value As List(Of CutToolPath))
				Me._innerToolPaths = value
			End Set
		End Property

		' Token: 0x17000013 RID: 19
		' (get) Token: 0x0600002D RID: 45 RVA: 0x00002408 File Offset: 0x00000608
		' (set) Token: 0x0600002E RID: 46 RVA: 0x00002410 File Offset: 0x00000610
		Public Property MillData As MillData
			Get
				Return Me._millData
			End Get
			Set(value As MillData)
				Me._millData = value
			End Set
		End Property

		' Token: 0x17000014 RID: 20
		' (get) Token: 0x0600002F RID: 47 RVA: 0x00002419 File Offset: 0x00000619
		' (set) Token: 0x06000030 RID: 48 RVA: 0x00002421 File Offset: 0x00000621
		Public Property FinalDepth As Double
			Get
				Return Me._finalDepth
			End Get
			Set(value As Double)
				Me._finalDepth = value
			End Set
		End Property

		' Token: 0x17000015 RID: 21
		' (get) Token: 0x06000031 RID: 49 RVA: 0x0000242A File Offset: 0x0000062A
		' (set) Token: 0x06000032 RID: 50 RVA: 0x00002432 File Offset: 0x00000632
		Public Property ToolFileName As String
			Get
				Return Me._toolFileName
			End Get
			Set(value As String)
				Me._toolFileName = value
			End Set
		End Property

		' Token: 0x17000016 RID: 22
		' (get) Token: 0x06000033 RID: 51 RVA: 0x0000243B File Offset: 0x0000063B
		' (set) Token: 0x06000034 RID: 52 RVA: 0x00002443 File Offset: 0x00000643
		Public Property IsDelete As Boolean
			Get
				Return Me._isDelete
			End Get
			Set(value As Boolean)
				Me._isDelete = value
			End Set
		End Property

		' Token: 0x17000017 RID: 23
		' (get) Token: 0x06000035 RID: 53 RVA: 0x0000244C File Offset: 0x0000064C
		' (set) Token: 0x06000036 RID: 54 RVA: 0x00002454 File Offset: 0x00000654
		Public Property Horizontal As Boolean
			Get
				Return Me._Horizontal
			End Get
			Set(value As Boolean)
				Me._Horizontal = value
			End Set
		End Property

		' Token: 0x17000018 RID: 24
		' (get) Token: 0x06000037 RID: 55 RVA: 0x0000245D File Offset: 0x0000065D
		' (set) Token: 0x06000038 RID: 56 RVA: 0x00002465 File Offset: 0x00000665
		Public Property DepthCutSpecial As Integer
			Get
				Return Me._depthCutSpecial
			End Get
			Set(value As Integer)
				Me._depthCutSpecial = value
			End Set
		End Property

		' Token: 0x17000019 RID: 25
		' (get) Token: 0x06000039 RID: 57 RVA: 0x0000246E File Offset: 0x0000066E
		' (set) Token: 0x0600003A RID: 58 RVA: 0x00002476 File Offset: 0x00000676
		Public Property CW As Boolean
			Get
				Return Me._cW
			End Get
			Set(value As Boolean)
				Me._cW = value
			End Set
		End Property

		' Token: 0x1700001A RID: 26
		' (get) Token: 0x0600003B RID: 59 RVA: 0x0000247F File Offset: 0x0000067F
		' (set) Token: 0x0600003C RID: 60 RVA: 0x00002487 File Offset: 0x00000687
		Public Property ToolGeo As Path
			Get
				Return Me._toolGeo
			End Get
			Set(value As Path)
				Me._toolGeo = value
			End Set
		End Property

		' Token: 0x1700001B RID: 27
		' (get) Token: 0x0600003D RID: 61 RVA: 0x00002490 File Offset: 0x00000690
		' (set) Token: 0x0600003E RID: 62 RVA: 0x00002498 File Offset: 0x00000698
		Public Property CutPriority As Integer
			Get
				Return Me._cutPriority
			End Get
			Set(value As Integer)
				Me._cutPriority = If((value > 4), 4, value)
			End Set
		End Property

		' Token: 0x1700001C RID: 28
		' (get) Token: 0x0600003F RID: 63 RVA: 0x000024A8 File Offset: 0x000006A8
		' (set) Token: 0x06000040 RID: 64 RVA: 0x000024B0 File Offset: 0x000006B0
		Public Property CutDirect As Integer
			Get
				Return Me._cutDirect
			End Get
			Set(value As Integer)
				Me._cutDirect = value
			End Set
		End Property

		' Token: 0x06000041 RID: 65 RVA: 0x00002D24 File Offset: 0x00000F24
		Public Sub DrawToolGeo(Drw As Drawing)
			Dim geo2D As Geo2D = Drw.Create2DGeometry(0.0, 0.0)
			Dim elements As Elements = Me.ToolPath.Elements
			Dim count As Integer = elements.Count
			For i As Integer = 1 To count
				Dim element As Element = elements.Item(i)
				If Not element.IsRapid Then
					geo2D = Drw.Create2DGeometry(element.StartXL, element.StartYL)
					IL_0061:
					For j As Integer = 1 To count
						Dim element2 As Element = elements.Item(j)
						If Not element2.IsRapid Then
							If element2.IsLine Then
								geo2D.AddLine(element2.EndXL, element2.EndYL)
							Else
								geo2D.AddArcPointCenter(element2.EndXL, element2.EndYL, element2.CenterXL, element2.CenterYL, element2.CW)
							End If
						End If
					Next
					Me.ToolGeo = geo2D.Finish()
					Me.ToolGeo.ToolInOut = AcamToolInOut.acamON_CENTER
					Me.ToolGeo.ToolSide = AcamToolSide.const_1
					Return
				End If
			Next
			GoTo IL_0061
		End Sub

		' Token: 0x06000042 RID: 66 RVA: 0x00002E24 File Offset: 0x00001024
		Public Sub RoughFinishBigPanel(Acam As App, Xin As Double, Yin As Double, <System.Runtime.InteropServices.OutAttribute()> ByRef Xout As Double, <System.Runtime.InteropServices.OutAttribute()> ByRef Yout As Double)
			Me.ToolGeo.SetStartPoint(Xin, Yin)
			If Me.MillData.FinalDepth <= -20F Then
				Dim num As Single = CSng((CDbl(Me.MillData.FinalDepth) * 0.6))
				Dim millData As MillData = Acam.CreateMillData()
				millData.SafeRapidLevel = Me.MillData.SafeRapidLevel
				millData.RapidDownTo = Me.MillData.RapidDownTo
				millData.SpindleSpeed = 24000F
				millData.CutFeed = 9000F
				millData.DownFeed = 2000F
				millData.FinalDepth = num
				Dim elements As Elements = Me.ToolGeo.Elements
				Dim millManualToolPath As MillManualToolPath = millData.ManualToolPath(Me.ToolGeo.GetFirstElem().StartXL, Me.ToolGeo.GetFirstElem().StartYL, CDbl(millData.FinalDepth))
				For i As Integer = 1 To elements.Count
					Dim element As Element = elements.Item(i)
					If element.IsLine Then
						millManualToolPath.imethod_0(element.EndXL, element.EndYL, CDbl(millData.FinalDepth))
					Else
						millManualToolPath.Add3DArcPointCenter(element.EndXL, element.EndYL, CDbl(millData.FinalDepth), element.CenterXL, element.CenterYL, element.CW)
					End If
				Next
				millManualToolPath.Finish()
				num = Me.MillData.FinalDepth
				millData.FinalDepth = num
				Dim millManualToolPath2 As MillManualToolPath = millData.ManualToolPath(Me.ToolGeo.GetFirstElem().StartXL, Me.ToolGeo.GetFirstElem().StartYL, CDbl(millData.FinalDepth))
				For j As Integer = 1 To elements.Count
					Dim element As Element = elements.Item(j)
					If element.IsLine Then
						millManualToolPath2.imethod_0(element.EndXL, element.EndYL, CDbl(millData.FinalDepth))
					Else
						millManualToolPath2.Add3DArcPointCenter(element.EndXL, element.EndYL, CDbl(millData.FinalDepth), element.CenterXL, element.CenterYL, element.CW)
					End If
				Next
				millManualToolPath2.Finish()
			Else
				Me.ToolGeo.Selected = True
				Me.MillData.RoughFinish()
				Me.ToolGeo.Selected = False
			End If
			Xout = Me.ToolGeo.GetLastElem().EndXL
			Yout = Me.ToolGeo.GetLastElem().EndYL
		End Sub

		' Token: 0x06000043 RID: 67 RVA: 0x00003074 File Offset: 0x00001274
		Public Sub RoughFinish(Acam As App, isSloopingLine As Boolean, backDist As Double, sloopingDist As Double, pointDist As Double)
			Me.SetStartPointOfSmallPanel(backDist)
			If isSloopingLine AndAlso sloopingDist > 0.0 Then
				If Me.MillData.FinalDepth <= -20F Then
					Dim num As Single = CSng((CDbl(Me.MillData.FinalDepth) * 0.6))
					Dim millData As MillData = Acam.CreateMillData()
					millData.SafeRapidLevel = Me.MillData.SafeRapidLevel
					millData.RapidDownTo = Me.MillData.RapidDownTo
					millData.SpindleSpeed = 24000F
					millData.CutFeed = 9000F
					millData.DownFeed = 2000F
					millData.FinalDepth = num
					Dim elements As Elements = Me.ToolGeo.Elements
					Dim millManualToolPath As MillManualToolPath = millData.ManualToolPath(Me.ToolGeo.GetFirstElem().StartXL, Me.ToolGeo.GetFirstElem().StartYL, CDbl(millData.FinalDepth))
					Dim element As Element
					For i As Integer = 1 To elements.Count
						element = elements.Item(i)
						If element.IsLine Then
							millManualToolPath.imethod_0(element.EndXL, element.EndYL, CDbl(millData.FinalDepth))
						Else
							millManualToolPath.Add3DArcPointCenter(element.EndXL, element.EndYL, CDbl(millData.FinalDepth), element.CenterXL, element.CenterYL, element.CW)
						End If
					Next
					millManualToolPath.Finish()
					num = Me.MillData.FinalDepth
					millData.FinalDepth = num
					Dim length As Double = Me.ToolGeo.Length
					Dim num2 As Double = length - sloopingDist
					Dim num3 As Double
					Dim num4 As Double
					Me.ToolGeo.PointAtDistanceAlongPathL(num2, num3, num4, element)
					Dim num5 As Integer = CInt(Math.Floor(sloopingDist / pointDist))
					Dim millManualToolPath2 As MillManualToolPath = millData.ManualToolPath(num3, num4, 0.0)
					Dim num6 As Double = CDbl((millData.FinalDepth / CSng(num5)))
					For j As Integer = 1 To num5
						If Me.ToolGeo.PointAtDistanceAlongPathL(num2 + pointDist * CDbl(j), num3, num4, element) Then
							millManualToolPath2.imethod_0(num3, num4, num6 * CDbl(j))
						End If
					Next
					millManualToolPath2.imethod_0(Me.ToolGeo.GetFirstElem().StartXL, Me.ToolGeo.GetFirstElem().StartYL, CDbl(millData.FinalDepth))
					For k As Integer = 1 To elements.Count
						element = elements.Item(k)
						If element.IsLine Then
							millManualToolPath2.imethod_0(element.EndXL, element.EndYL, CDbl(millData.FinalDepth))
						Else
							millManualToolPath2.Add3DArcPointCenter(element.EndXL, element.EndYL, CDbl(millData.FinalDepth), element.CenterXL, element.CenterYL, element.CW)
						End If
					Next
					millManualToolPath2.Finish()
					Return
				End If
				Dim finalDepth As Single = Me.MillData.FinalDepth
				If Me.MillData.DepthsOfCutSpecified Then
					Dim millData2 As MillData = Acam.CreateMillData()
					millData2.SafeRapidLevel = Me.MillData.SafeRapidLevel
					millData2.RapidDownTo = Me.MillData.RapidDownTo
					millData2.SpindleSpeed = 24000F
					millData2.CutFeed = 9000F
					millData2.DownFeed = 2000F
					millData2.FinalDepth = finalDepth / 2F
					Dim elements2 As Elements = Me.ToolGeo.Elements
					Dim millManualToolPath3 As MillManualToolPath = millData2.ManualToolPath(Me.ToolGeo.GetFirstElem().StartXL, Me.ToolGeo.GetFirstElem().StartYL, CDbl(millData2.FinalDepth))
					For l As Integer = 1 To elements2.Count
						Dim element2 As Element = elements2.Item(l)
						If element2.IsLine Then
							millManualToolPath3.imethod_0(element2.EndXL, element2.EndYL, CDbl(millData2.FinalDepth))
						Else
							millManualToolPath3.Add3DArcPointCenter(element2.EndXL, element2.EndYL, CDbl(millData2.FinalDepth), element2.CenterXL, element2.CenterYL, element2.CW)
						End If
					Next
					millManualToolPath3.Finish()
				End If
				Dim millData3 As MillData = Acam.CreateMillData()
				millData3.SafeRapidLevel = Me.MillData.SafeRapidLevel
				millData3.RapidDownTo = Me.MillData.RapidDownTo
				millData3.SpindleSpeed = 24000F
				millData3.CutFeed = 9000F
				millData3.DownFeed = 2000F
				millData3.FinalDepth = finalDepth
				Dim elements3 As Elements = Me.ToolGeo.Elements
				Dim length2 As Double = Me.ToolGeo.Length
				Dim num7 As Double = length2 - sloopingDist
				Dim num8 As Double
				Dim num9 As Double
				Dim element3 As Element
				Me.ToolGeo.PointAtDistanceAlongPathL(num7, num8, num9, element3)
				Dim num10 As Integer = CInt(Math.Floor(sloopingDist / pointDist))
				Dim millManualToolPath4 As MillManualToolPath = millData3.ManualToolPath(num8, num9, 0.0)
				Dim num11 As Double = CDbl((millData3.FinalDepth / CSng(num10)))
				For m As Integer = 1 To num10
					If Me.ToolGeo.PointAtDistanceAlongPathL(num7 + pointDist * CDbl(m), num8, num9, element3) Then
						millManualToolPath4.imethod_0(num8, num9, num11 * CDbl(m))
					End If
				Next
				millManualToolPath4.imethod_0(Me.ToolGeo.GetFirstElem().StartXL, Me.ToolGeo.GetFirstElem().StartYL, CDbl(millData3.FinalDepth))
				For n As Integer = 1 To elements3.Count
					element3 = elements3.Item(n)
					If element3.IsLine Then
						millManualToolPath4.imethod_0(element3.EndXL, element3.EndYL, CDbl(millData3.FinalDepth))
					Else
						millManualToolPath4.Add3DArcPointCenter(element3.EndXL, element3.EndYL, CDbl(millData3.FinalDepth), element3.CenterXL, element3.CenterYL, element3.CW)
					End If
				Next
				millManualToolPath4.Finish()
				Return
			Else
				If Me.MillData.FinalDepth <= -20F Then
					Dim num12 As Single = CSng((CDbl(Me.MillData.FinalDepth) * 0.6))
					Dim millData4 As MillData = Acam.CreateMillData()
					millData4.SafeRapidLevel = Me.MillData.SafeRapidLevel
					millData4.RapidDownTo = Me.MillData.RapidDownTo
					millData4.SpindleSpeed = 24000F
					millData4.CutFeed = 9000F
					millData4.DownFeed = 2000F
					millData4.FinalDepth = num12
					Dim elements4 As Elements = Me.ToolGeo.Elements
					Dim millManualToolPath5 As MillManualToolPath = millData4.ManualToolPath(Me.ToolGeo.GetFirstElem().StartXL, Me.ToolGeo.GetFirstElem().StartYL, CDbl(num12))
					For num13 As Integer = 1 To elements4.Count
						Dim element4 As Element = elements4.Item(num13)
						If element4.IsLine Then
							millManualToolPath5.imethod_0(element4.EndXL, element4.EndYL, CDbl(millData4.FinalDepth))
						Else
							millManualToolPath5.Add3DArcPointCenter(element4.EndXL, element4.EndYL, CDbl(millData4.FinalDepth), element4.CenterXL, element4.CenterYL, element4.CW)
						End If
					Next
					millManualToolPath5.Finish()
					num12 = Me.MillData.FinalDepth
					millData4.FinalDepth = num12
					Dim millManualToolPath6 As MillManualToolPath = millData4.ManualToolPath(Me.ToolGeo.GetFirstElem().StartXL, Me.ToolGeo.GetFirstElem().StartYL, CDbl(num12))
					For num14 As Integer = 1 To elements4.Count
						Dim element4 As Element = elements4.Item(num14)
						If element4.IsLine Then
							millManualToolPath6.imethod_0(element4.EndXL, element4.EndYL, CDbl(millData4.FinalDepth))
						Else
							millManualToolPath6.Add3DArcPointCenter(element4.EndXL, element4.EndYL, CDbl(millData4.FinalDepth), element4.CenterXL, element4.CenterYL, element4.CW)
						End If
					Next
					millManualToolPath6.Finish()
					Return
				End If
				Me.ToolGeo.Selected = True
				Me.MillData.RoughFinish()
				Me.ToolGeo.Selected = False
				Return
			End If
		End Sub

		' Token: 0x06000044 RID: 68 RVA: 0x00003834 File Offset: 0x00001A34
		Private Sub SetStartPointOfSmallPanel(backValue As Double)
			If Me.CutDirect <> 15 AndAlso Me.CutDirect <> 0 Then
				If Me.CutDirect = 14 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX, Me.MaxY - backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinX, Me.MinY + backValue)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX, Me.MaxY - backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinX, Me.MinY + backValue)
						Return
					End If
				ElseIf Me.CutDirect = 13 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MaxY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinY + backValue, Me.MaxY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MaxY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinY + backValue, Me.MaxY)
						Return
					End If
				ElseIf Me.CutDirect = 12 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MaxY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinY + backValue, Me.MaxY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX, Me.MaxY - backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinX, Me.MinY + backValue)
						Return
					End If
				ElseIf Me.CutDirect = 11 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
						Return
					End If
				ElseIf Me.CutDirect = 10 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
						Return
					End If
				ElseIf Me.CutDirect = 9 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MaxY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MaxY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
						Return
					End If
				ElseIf Me.CutDirect = 8 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MaxY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MaxY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
						Return
					End If
				ElseIf Me.CutDirect = 7 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MinY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MinY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MinY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MinY)
						Return
					End If
				ElseIf Me.CutDirect = 6 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MinY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MinY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX, Me.MaxY - backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinX, Me.MinY + backValue)
						Return
					End If
				ElseIf Me.CutDirect = 5 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MinY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MinY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinX, Me.MaxY - backValue)
						Return
					End If
				ElseIf Me.CutDirect = 4 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MinY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MinY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX, Me.MaxY - backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MinX, Me.MinY + backValue)
						Return
					End If
				ElseIf Me.CutDirect = 3 Then
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MinY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MinY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
						Return
					End If
				Else
					If Me.CutDirect <> 2 Then
						If Me.CutDirect = 1 Then
							If Me.Horizontal Then
								If Me.CW Then
									Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MaxY)
									Return
								End If
								Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MaxY)
								Return
							Else
								If Me.CW Then
									Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
									Return
								End If
								Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
							End If
						End If
						Return
					End If
					If Me.Horizontal Then
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MinY)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MinY)
						Return
					Else
						If Me.CW Then
							Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
							Return
						End If
						Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
						Return
					End If
				End If
			ElseIf Me.Horizontal Then
				If Me.CW Then
					Me.ToolGeo.SetStartPoint(Me.MaxX - backValue, Me.MaxY)
					Return
				End If
				Me.ToolGeo.SetStartPoint(Me.MinX + backValue, Me.MaxY)
				Return
			Else
				If Me.CW Then
					Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MinY + backValue)
					Return
				End If
				Me.ToolGeo.SetStartPoint(Me.MaxX, Me.MaxY - backValue)
				Return
			End If
		End Sub

		' Token: 0x06000045 RID: 69 RVA: 0x00002395 File Offset: 0x00000595
		Public Sub New()
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
		End Sub

		' Token: 0x0400000E RID: 14
		Private _minX As Double

		' Token: 0x0400000F RID: 15
		Private _maxX As Double

		' Token: 0x04000010 RID: 16
		Private _minY As Double

		' Token: 0x04000011 RID: 17
		Private _maxY As Double

		' Token: 0x04000012 RID: 18
		Private toolPath As Path

		' Token: 0x04000013 RID: 19
		Private _innerToolPaths As List(Of CutToolPath)

		' Token: 0x04000014 RID: 20
		Private _millData As MillData

		' Token: 0x04000015 RID: 21
		Private _finalDepth As Double

		' Token: 0x04000016 RID: 22
		Private _toolFileName As String

		' Token: 0x04000017 RID: 23
		Private _isDelete As Boolean

		' Token: 0x04000018 RID: 24
		Private _Horizontal As Boolean

		' Token: 0x04000019 RID: 25
		Private _depthCutSpecial As Integer

		' Token: 0x0400001A RID: 26
		Private _cW As Boolean

		' Token: 0x0400001B RID: 27
		Private _toolGeo As Path

		' Token: 0x0400001C RID: 28
		Private _cutPriority As Integer

		' Token: 0x0400001D RID: 29
		Private _cutDirect As Integer
	End Class
End Namespace
