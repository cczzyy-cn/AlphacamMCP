Imports System
Imports System.Collections
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x0200002D RID: 45
	<CompilerGenerated()>
	<TypeIdentifier()>
	<DefaultMember("Item")>
	<Guid("31C27FD6-AA56-11D2-991A-00104B4AF281")>
	<ComImport()>
	Public Interface IOperations
		Inherits IEnumerable

		' Token: 0x06000153 RID: 339
		Sub _VtblGap1_1()

		' Token: 0x06000154 RID: 340
		<DispId(0)>
		Function Item(<[In]()> Index As Integer) As <MarshalAs(UnmanagedType.[Interface])> Operation

		' Token: 0x17000061 RID: 97
		' (get) Token: 0x06000155 RID: 341
		ReadOnly Property Count As Integer
	End Interface
End Namespace
