Imports System
Imports System.Collections
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x02000030 RID: 48
	<TypeIdentifier()>
	<Guid("31C27FD4-AA56-11D2-991A-00104B4AF281")>
	<CompilerGenerated()>
	<DefaultMember("Item")>
	<ComImport()>
	Public Interface ISubOperations
		Inherits IEnumerable

		' Token: 0x06000158 RID: 344
		Sub _VtblGap1_1()

		' Token: 0x06000159 RID: 345
		<DispId(0)>
		Function Item(<[In]()> Index As Integer) As <MarshalAs(UnmanagedType.[Interface])> SubOperation

		' Token: 0x17000063 RID: 99
		' (get) Token: 0x0600015A RID: 346
		ReadOnly Property Count As Integer
	End Interface
End Namespace
