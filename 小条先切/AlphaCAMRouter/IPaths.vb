Imports System
Imports System.Collections
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x0200002B RID: 43
	<DefaultMember("Item")>
	<CompilerGenerated()>
	<Guid("AFA20680-8305-11D2-98D1-00104B4AF281")>
	<TypeIdentifier()>
	<ComImport()>
	Public Interface IPaths
		Inherits IEnumerable

		' Token: 0x06000150 RID: 336
		Sub _VtblGap1_1()

		' Token: 0x06000151 RID: 337
		<DispId(0)>
		Function Item(<[In]()> Index As Integer) As <MarshalAs(UnmanagedType.[Interface])> Path

		' Token: 0x17000060 RID: 96
		' (get) Token: 0x06000152 RID: 338
		ReadOnly Property Count As Integer
	End Interface
End Namespace
