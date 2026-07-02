Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x0200001D RID: 29
	<CompilerGenerated()>
	<Guid("995E5092-462D-11D2-9869-00104B4AF281")>
	<TypeIdentifier()>
	<ComImport()>
	Public Interface IPath
		' Token: 0x17000041 RID: 65
		' (get) Token: 0x060000E9 RID: 233
		' (set) Token: 0x060000EA RID: 234
		Property Selected As Boolean

		' Token: 0x17000042 RID: 66
		' (get) Token: 0x060000EB RID: 235
		' (set) Token: 0x060000EC RID: 236
		Property ToolSide As AcamToolSide

		' Token: 0x17000043 RID: 67
		' (get) Token: 0x060000ED RID: 237
		ReadOnly Property Closed As Boolean

		' Token: 0x17000044 RID: 68
		' (get) Token: 0x060000EE RID: 238
		' (set) Token: 0x060000EF RID: 239
		Property ToolInOut As AcamToolInOut

		' Token: 0x17000045 RID: 69
		' (get) Token: 0x060000F0 RID: 240
		' (set) Token: 0x060000F1 RID: 241
		Property CW As Short

		' Token: 0x17000046 RID: 70
		' (get) Token: 0x060000F2 RID: 242
		ReadOnly Property MaxXL As Double

		' Token: 0x17000047 RID: 71
		' (get) Token: 0x060000F3 RID: 243
		ReadOnly Property MinXL As Double

		' Token: 0x17000048 RID: 72
		' (get) Token: 0x060000F4 RID: 244
		ReadOnly Property MaxYL As Double

		' Token: 0x17000049 RID: 73
		' (get) Token: 0x060000F5 RID: 245
		ReadOnly Property MinYL As Double

		' Token: 0x060000F6 RID: 246
		Sub _VtblGap1_2()

		' Token: 0x1700004A RID: 74
		' (get) Token: 0x060000F7 RID: 247
		ReadOnly Property Length As Double

		' Token: 0x060000F8 RID: 248
		Sub _VtblGap2_10()

		' Token: 0x1700004B RID: 75
		' (get) Token: 0x060000F9 RID: 249
		' (set) Token: 0x060000FA RID: 250
		Property Visible As Boolean

		' Token: 0x060000FB RID: 251
		Sub _VtblGap3_6()

		' Token: 0x060000FC RID: 252
		<DispId(12)>
		Sub SetStartPoint(<[In]()> X As Double, <[In]()> Y As Double)

		' Token: 0x060000FD RID: 253
		Sub _VtblGap4_2()

		' Token: 0x060000FE RID: 254
		<DispId(15)>
		Function GetFirstElem() As <MarshalAs(UnmanagedType.[Interface])> Element

		' Token: 0x060000FF RID: 255
		<DispId(16)>
		Function GetLastElem() As <MarshalAs(UnmanagedType.[Interface])> Element

		' Token: 0x06000100 RID: 256
		Sub _VtblGap5_5()

		' Token: 0x06000101 RID: 257
		<DispId(22)>
		Function Copy() As <MarshalAs(UnmanagedType.[Interface])> Path

		' Token: 0x06000102 RID: 258
		<DispId(23)>
		Sub MoveL(<[In]()> X As Double, <[In]()> Y As Double)

		' Token: 0x06000103 RID: 259
		Sub _VtblGap6_5()

		' Token: 0x06000104 RID: 260
		<DispId(29)>
		Sub ScaleL2(<[In]()> FactorX As Double, <[In]()> FactorY As Double, <[In]()> BaseX As Double, <[In]()> BaseY As Double)

		' Token: 0x06000105 RID: 261
		Sub _VtblGap7_12()

		' Token: 0x06000106 RID: 262
		<DispId(42)>
		Function PointAtDistanceAlongPathL(<[In]()> Distance As Double, <System.Runtime.InteropServices.OutAttribute()> ByRef XP As Double, <System.Runtime.InteropServices.OutAttribute()> ByRef YP As Double, <MarshalAs(UnmanagedType.[Interface])> <System.Runtime.InteropServices.OutAttribute()> ByRef Element As Element) As Boolean

		' Token: 0x06000107 RID: 263
		<DispId(43)>
		Sub Reverse()

		' Token: 0x06000108 RID: 264
		Sub _VtblGap8_2()

		' Token: 0x06000109 RID: 265
		<DispId(46)>
		Function GetMillData() As <MarshalAs(UnmanagedType.[Interface])> MillData

		' Token: 0x0600010A RID: 266
		Sub _VtblGap9_7()

		' Token: 0x0600010B RID: 267
		<DispId(54)>
		Function TestIntersectPath(<MarshalAs(UnmanagedType.[Interface])> <[In]()> Path2 As Path, <[In]()> XShift2 As Double, <[In]()> YShift2 As Double) As Boolean

		' Token: 0x0600010C RID: 268
		Sub _VtblGap10_20()

		' Token: 0x0600010D RID: 269
		<DispId(74)>
		Function GetFeedExtent(<System.Runtime.InteropServices.OutAttribute()> ByRef MinX As Double, <System.Runtime.InteropServices.OutAttribute()> ByRef MinY As Double, <System.Runtime.InteropServices.OutAttribute()> ByRef MaxX As Double, <System.Runtime.InteropServices.OutAttribute()> ByRef MaxY As Double) As Boolean

		' Token: 0x0600010E RID: 270
		Sub _VtblGap11_2()

		' Token: 0x1700004C RID: 76
		' (get) Token: 0x0600010F RID: 271
		ReadOnly Property Elements As Elements

		' Token: 0x06000110 RID: 272
		Sub _VtblGap12_32()

		' Token: 0x1700004D RID: 77
		' (get) Token: 0x06000111 RID: 273
		ReadOnly Property Is3D As Boolean

		' Token: 0x06000112 RID: 274
		Sub _VtblGap13_32()

		' Token: 0x06000113 RID: 275
		<DispId(127)>
		Sub Delete()
	End Interface
End Namespace
