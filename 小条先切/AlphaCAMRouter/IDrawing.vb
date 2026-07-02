Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x0200001B RID: 27
	<TypeIdentifier()>
	<CompilerGenerated()>
	<Guid("1A172592-4565-11D2-9866-00104B4AF281")>
	<ComImport()>
	Public Interface IDrawing
		' Token: 0x060000DC RID: 220
		Sub _VtblGap1_13()

		' Token: 0x060000DD RID: 221
		<DispId(9)>
		Sub ZoomAll()

		' Token: 0x060000DE RID: 222
		Sub _VtblGap2_18()

		' Token: 0x060000DF RID: 223
		<DispId(28)>
		Function Create2DGeometry(<[In]()> X As Double, <[In]()> Y As Double) As <MarshalAs(UnmanagedType.[Interface])> Geo2D

		' Token: 0x060000E0 RID: 224
		Sub _VtblGap3_15()

		' Token: 0x060000E1 RID: 225
		<DispId(44)>
		Function Join() As <MarshalAs(UnmanagedType.[Interface])> Paths

		' Token: 0x060000E2 RID: 226
		Sub _VtblGap4_5()

		' Token: 0x060000E3 RID: 227
		<DispId(50)>
		Function Create2DLine(<[In]()> X1 As Double, <[In]()> Y1 As Double, <[In]()> X2 As Double, <[In]()> Y2 As Double) As <MarshalAs(UnmanagedType.[Interface])> Path

		' Token: 0x060000E4 RID: 228
		Sub _VtblGap5_42()

		' Token: 0x060000E5 RID: 229
		<DispId(93)>
		Function GetNestInformation() As <MarshalAs(UnmanagedType.IDispatch)> Object

		' Token: 0x17000040 RID: 64
		' (get) Token: 0x060000E6 RID: 230
		ReadOnly Property Operations As Operations

		' Token: 0x060000E7 RID: 231
		Sub _VtblGap6_194()

		' Token: 0x060000E8 RID: 232
		<DispId(269)>
		Sub OrderToolPathsInNestedSheets()
	End Interface
End Namespace
