Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x02000026 RID: 38
	<CompilerGenerated()>
	<TypeIdentifier()>
	<Guid("88A2FF74-4309-11D2-9861-00104B4AF281")>
	<ComImport()>
	Public Interface IAlphaCamApp
		' Token: 0x06000144 RID: 324
		Sub _VtblGap1_2()

		' Token: 0x1700005F RID: 95
		' (get) Token: 0x06000145 RID: 325
		ReadOnly Property ActiveDrawing As Drawing

		' Token: 0x06000146 RID: 326
		Sub _VtblGap2_15()

		' Token: 0x06000147 RID: 327
		<DispId(18)>
		Function CreateMillData() As <MarshalAs(UnmanagedType.[Interface])> MillData

		' Token: 0x06000148 RID: 328
		Sub _VtblGap3_1()

		' Token: 0x06000149 RID: 329
		<DispId(20)>
		Function SelectTool(<MarshalAs(UnmanagedType.BStr)> <[In]()> Name As String) As <MarshalAs(UnmanagedType.[Interface])> MillTool
	End Interface
End Namespace
