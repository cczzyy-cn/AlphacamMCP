Imports System
Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Globalization
Imports System.Resources
Imports System.Runtime.CompilerServices

Namespace 小条先切.Properties
	' Token: 0x02000014 RID: 20
	<CompilerGenerated()>
	<DebuggerNonUserCode()>
	<GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")>
	Friend Class Resources
		' Token: 0x060000D6 RID: 214 RVA: 0x00002395 File Offset: 0x00000595
		Friend Sub New()
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
		End Sub

		' Token: 0x1700003D RID: 61
		' (get) Token: 0x060000D7 RID: 215 RVA: 0x00007F04 File Offset: 0x00006104
		<EditorBrowsable(EditorBrowsableState.Advanced)>
		Friend Shared ReadOnly Property ResourceManager As ResourceManager
			Get
				If Object.ReferenceEquals(Resources.resourceMan, Nothing) Then
					Dim resourceManager As ResourceManager = New ResourceManager("小条先切.Properties.Resources", GetType(Resources).Assembly)
					Resources.resourceMan = resourceManager
				End If
				Return Resources.resourceMan
			End Get
		End Property

		' Token: 0x1700003E RID: 62
		' (set) Token: 0x060000D8 RID: 216 RVA: 0x0000285F File Offset: 0x00000A5F
		<EditorBrowsable(EditorBrowsableState.Advanced)>
		Friend Shared WriteOnly Property Culture As CultureInfo
			Set(value As CultureInfo)
				Resources.resourceCulture = value
			End Set
		End Property

		' Token: 0x04000069 RID: 105
		Private Shared resourceMan As ResourceManager

		' Token: 0x0400006A RID: 106
		Private Shared resourceCulture As CultureInfo
	End Class
End Namespace
