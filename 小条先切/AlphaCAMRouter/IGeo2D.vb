Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x02000022 RID: 34
	<Guid("2EBF8F62-48A8-11D2-9872-00104B4AF281")>
	<TypeIdentifier()>
	<CompilerGenerated()>
	<ComImport()>
	Public Interface IGeo2D
		' Token: 0x0600012C RID: 300
		<DispId(1)>
		Sub AddLine(<[In]()> X As Double, <[In]()> Y As Double)

		' Token: 0x0600012D RID: 301
		<DispId(2)>
		Function Finish() As <MarshalAs(UnmanagedType.[Interface])> Path

		' Token: 0x0600012E RID: 302
		Sub _VtblGap1_3()

		' Token: 0x0600012F RID: 303
		<DispId(6)>
		Sub AddArcPointCenter(<[In]()> X As Double, <[In]()> Y As Double, <[In]()> XCen As Double, <[In]()> YCen As Double, <[In]()> CW As Boolean)
	End Interface
End Namespace
