Imports System
Imports System.Collections
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x0200001F RID: 31
	<TypeIdentifier()>
	<CompilerGenerated()>
	<Guid("9F4323F0-A87B-11D2-9917-00104B4AF281")>
	<DefaultMember("Item")>
	<ComImport()>
	Public Interface IElements
		Inherits IEnumerable

		' Token: 0x06000114 RID: 276
		Sub _VtblGap1_1()

		' Token: 0x06000115 RID: 277
		<DispId(0)>
		Function Item(<[In]()> Index As Integer) As <MarshalAs(UnmanagedType.[Interface])> Element

		' Token: 0x1700004E RID: 78
		' (get) Token: 0x06000116 RID: 278
		ReadOnly Property Count As Integer
	End Interface
End Namespace
