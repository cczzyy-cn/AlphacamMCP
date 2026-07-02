Imports System
Imports System.Reflection

' Token: 0x02000035 RID: 53
Friend Class Class1
	' Token: 0x06000164 RID: 356 RVA: 0x00007F44 File Offset: 0x00006144
	Friend Shared Sub B3Pn1h99atayT(typemdt As Integer)
		Dim type As Type = Class1.eYhbKmqb0.ResolveType(33554432 + typemdt)
		For Each fieldInfo As FieldInfo In type.GetFields()
			Dim methodInfo As MethodInfo = CType(Class1.eYhbKmqb0.ResolveMethod(fieldInfo.MetadataToken + 100663296), MethodInfo)
			fieldInfo.SetValue(Nothing, CType([Delegate].CreateDelegate(type, methodInfo), MulticastDelegate))
		Next
	End Sub

	' Token: 0x06000165 RID: 357 RVA: 0x00002395 File Offset: 0x00000595
	Public Sub New()
		Class2.C5hn1h9zVegKm()
		MyBase..ctor()
	End Sub

	' Token: 0x06000166 RID: 358 RVA: 0x00002896 File Offset: 0x00000A96
	' Note: this type is marked as 'beforefieldinit'.
	Shared Sub New()
		Class2.C5hn1h9zVegKm()
		Class1.eYhbKmqb0 = GetType(Class1).Assembly.ManifestModule
	End Sub

	' Token: 0x04000074 RID: 116
	Friend Shared eYhbKmqb0 As [Module]

	' Token: 0x02000036 RID: 54
	' (Invoke) Token: 0x06000168 RID: 360
	Friend Delegate Sub Delegate0(o As Object)
End Class
