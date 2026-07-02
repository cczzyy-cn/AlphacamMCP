Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x02000028 RID: 40
	<CompilerGenerated()>
	<TypeIdentifier()>
	<Guid("B8038580-A9A6-11D2-9919-00104B4AF281")>
	<ComImport()>
	Public Interface GInterface0
		' Token: 0x0600014A RID: 330
		Sub _VtblGap1_1()

		' Token: 0x0600014B RID: 331
		<DispId(2)>
		Sub imethod_0(<[In]()> X As Double, <[In]()> Y As Double, <[In]()> Z As Double)

		' Token: 0x0600014C RID: 332
		Sub _VtblGap2_9()

		' Token: 0x0600014D RID: 333
		<DispId(12)>
		Sub Add3DArcPointCenter(<[In]()> X As Double, <[In]()> Y As Double, <[In]()> Z As Double, <[In]()> XCen As Double, <[In]()> YCen As Double, <[In]()> CW As Boolean)

		' Token: 0x0600014E RID: 334
		Sub _VtblGap3_2()

		' Token: 0x0600014F RID: 335
		<DispId(15)>
		Function Finish() As <MarshalAs(UnmanagedType.[Interface])> Paths
	End Interface
End Namespace
