Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Threading
Imports AlphaCAMRouter
Imports Microsoft.CSharp.RuntimeBinder

Namespace 小条先切
	' Token: 0x0200000C RID: 12
	Public Module AlphaTool
		' Token: 0x06000091 RID: 145 RVA: 0x0000609C File Offset: 0x0000429C
		Public Function GetAlphaApp() As App
			Dim processesByName As Process() = Process.GetProcessesByName("Acam")
			If processesByName.Length <= 0 Then
				Return Nothing
			End If
			Return CType(Marshal.GetActiveObject("aroutaps.application"), App)
		End Function

		' Token: 0x06000092 RID: 146 RVA: 0x000060CC File Offset: 0x000042CC
		Public Function HasNesting(Acam As App) As Boolean
			Dim activeDrawing As Drawing = Acam.ActiveDrawing
			Dim flag As Boolean = False
			Try
				activeDrawing.GetNestInformation()
				flag = True
			Catch
			Finally
			End Try
			Return flag
		End Function

		' Token: 0x06000093 RID: 147 RVA: 0x00006114 File Offset: 0x00004314
		Public Function GetDrwToolName(Acam As App) As List(Of String)
			Dim list As List(Of String) = New List(Of String)()
			Dim activeDrawing As Drawing = Acam.ActiveDrawing
			activeDrawing.ZoomAll()
			Dim operations As Operations = activeDrawing.Operations
			Dim count As Integer = operations.Count
			For i As Integer = 1 To count
				Dim operation As Operation = operations.Item(i)
				Dim subOperations As SubOperations = operation.SubOperations
				Dim count2 As Integer = subOperations.Count
				For j As Integer = 1 To count2
					Dim subOperation As SubOperation = subOperations.Item(j)
					list.Add(subOperation.Name)
				Next
			Next
			Return list.Distinct().ToList()
		End Function

		' Token: 0x06000094 RID: 148 RVA: 0x000061A4 File Offset: 0x000043A4
		Public Function GetDrwNesting(Acam As App) As List(Of DrwNest)
			Dim list As List(Of DrwNest) = New List(Of DrwNest)()
			Dim activeDrawing As Drawing = Acam.ActiveDrawing
			Dim nestInformation As Object = activeDrawing.GetNestInformation()
			If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site1 Is Nothing Then
				AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site1 = CallSite(Of Func(Of CallSite, Object, Object)).Create(Binder.GetMember(CSharpBinderFlags.None, "Sheets", GetType(AlphaTool), New CSharpArgumentInfo() { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, Nothing) }))
			End If
			Dim obj As Object = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site1.Target(AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site1, nestInformation)
			If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site2 Is Nothing Then
				AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site2 = CallSite(Of Func(Of CallSite, Object, Integer)).Create(Binder.Convert(CSharpBinderFlags.None, GetType(Integer), GetType(AlphaTool)))
			End If
			Dim target As Func(Of CallSite, Object, Integer) = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site2.Target
			Dim <>p__Site As CallSite = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site2
			If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site3 Is Nothing Then
				AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site3 = CallSite(Of Func(Of CallSite, Object, Object)).Create(Binder.GetMember(CSharpBinderFlags.None, "Count", GetType(AlphaTool), New CSharpArgumentInfo() { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, Nothing) }))
			End If
			Dim num As Integer = target(<>p__Site, AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site3.Target(AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site3, obj))
			For i As Integer = 1 To num
				If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site4 Is Nothing Then
					AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site4 = CallSite(Of Func(Of CallSite, Object, Integer, Object)).Create(Binder.InvokeMember(CSharpBinderFlags.None, "Item", Nothing, GetType(AlphaTool), New CSharpArgumentInfo() { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, Nothing), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, Nothing) }))
				End If
				Dim obj2 As Object = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site4.Target(AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site4, obj, i)
				Dim drwNest As DrwNest = New DrwNest()
				Dim drwNest2 As DrwNest = drwNest
				If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site5 Is Nothing Then
					AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site5 = CallSite(Of Func(Of CallSite, Object, String)).Create(Binder.Convert(CSharpBinderFlags.None, GetType(String), GetType(AlphaTool)))
				End If
				Dim target2 As Func(Of CallSite, Object, String) = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site5.Target
				Dim <>p__Site2 As CallSite = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site5
				If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site6 Is Nothing Then
					AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site6 = CallSite(Of Func(Of CallSite, Object, Object)).Create(Binder.GetMember(CSharpBinderFlags.None, "Name", GetType(AlphaTool), New CSharpArgumentInfo() { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, Nothing) }))
				End If
				drwNest2.NestName = target2(<>p__Site2, AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site6.Target(AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site6, obj2))
				Dim num2 As Integer = drwNest.NestName.IndexOf(" "c) + 2
				drwNest.SheetNo = Convert.ToInt32(drwNest.NestName.Substring(num2, drwNest.NestName.Length - num2))
				Dim drwNest3 As DrwNest = drwNest
				If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site7 Is Nothing Then
					AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site7 = CallSite(Of Func(Of CallSite, Object, Double)).Create(Binder.Convert(CSharpBinderFlags.None, GetType(Double), GetType(AlphaTool)))
				End If
				Dim target3 As Func(Of CallSite, Object, Double) = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site7.Target
				Dim <>p__Site3 As CallSite = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site7
				If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site8 Is Nothing Then
					AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site8 = CallSite(Of Func(Of CallSite, Object, Object)).Create(Binder.GetMember(CSharpBinderFlags.None, "Thickness", GetType(AlphaTool), New CSharpArgumentInfo() { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, Nothing) }))
				End If
				drwNest3.Thickness = target3(<>p__Site3, AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site8.Target(AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site8, obj2))
				If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site9 Is Nothing Then
					AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site9 = CallSite(Of Func(Of CallSite, Object, Path)).Create(Binder.Convert(CSharpBinderFlags.None, GetType(Path), GetType(AlphaTool)))
				End If
				Dim target4 As Func(Of CallSite, Object, Path) = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site9.Target
				Dim <>p__Site4 As CallSite = AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Site9
				If AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Sitea Is Nothing Then
					AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Sitea = CallSite(Of Func(Of CallSite, Object, Object)).Create(Binder.GetMember(CSharpBinderFlags.None, "Geometry", GetType(AlphaTool), New CSharpArgumentInfo() { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, Nothing) }))
				End If
				Dim path As Path = target4(<>p__Site4, AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Sitea.Target(AlphaTool.<GetDrwNesting>o__SiteContainer0.<>p__Sitea, obj2))
				drwNest.MinX = path.MinXL
				drwNest.MaxX = path.MaxXL
				drwNest.MinY = path.MinYL
				drwNest.MaxY = path.MaxYL
				list.Add(drwNest)
			Next
			Return list
		End Function

		' Token: 0x06000095 RID: 149 RVA: 0x00006564 File Offset: 0x00004764
		Public Function GetToolGeoByToolName(Drw As Drawing, toolName As String, bool_0 As Boolean, includeOpen As Boolean) As List(Of CutToolPath)
			Dim list As List(Of CutToolPath) = New List(Of CutToolPath)()
			Dim operations As Operations = Drw.Operations
			Dim count As Integer = operations.Count
			Dim list2 As List(Of Thread) = New List(Of Thread)()
			Dim list3 As List(Of OpsPara) = New List(Of OpsPara)()
			For i As Integer = 1 To count
				Dim operation As Operation = operations.Item(i)
				Dim opsPara As OpsPara = New OpsPara(bool_0, includeOpen, toolName, operation)
				list3.Add(opsPara)
				Dim thread As Thread = AddressOf opsPara.[Set]
				list2.Add(thread)
				thread.Start()
			Next
			While True
				IL_009D:
				Dim flag As Boolean = False
				Thread.CurrentThread.Join(1000)
				Dim j As Integer = 0
				While j < list2.Count
					If Not list2(j).IsAlive Then
						j += 1
					Else
						flag = True
						IL_0099:
						If flag Then
							GoTo IL_009D
						End If
						GoTo IL_00B5
					End If
				End While
				GoTo IL_0099
			End While
			IL_00B5:
			For k As Integer = 0 To list3.Count - 1
				list2(k).Abort()
				list.AddRange(list3(k).Tps)
			Next
			Return list
		End Function

		' Token: 0x0200000D RID: 13
		<CompilerGenerated()>
		Private NotInheritable Class <GetDrwNesting>o__SiteContainer0
			' Token: 0x04000039 RID: 57
			Public Shared <>p__Site1 As CallSite(Of Func(Of CallSite, Object, Object))

			' Token: 0x0400003A RID: 58
			Public Shared <>p__Site2 As CallSite(Of Func(Of CallSite, Object, Integer))

			' Token: 0x0400003B RID: 59
			Public Shared <>p__Site3 As CallSite(Of Func(Of CallSite, Object, Object))

			' Token: 0x0400003C RID: 60
			Public Shared <>p__Site4 As CallSite(Of Func(Of CallSite, Object, Integer, Object))

			' Token: 0x0400003D RID: 61
			Public Shared <>p__Site5 As CallSite(Of Func(Of CallSite, Object, String))

			' Token: 0x0400003E RID: 62
			Public Shared <>p__Site6 As CallSite(Of Func(Of CallSite, Object, Object))

			' Token: 0x0400003F RID: 63
			Public Shared <>p__Site7 As CallSite(Of Func(Of CallSite, Object, Double))

			' Token: 0x04000040 RID: 64
			Public Shared <>p__Site8 As CallSite(Of Func(Of CallSite, Object, Object))

			' Token: 0x04000041 RID: 65
			Public Shared <>p__Site9 As CallSite(Of Func(Of CallSite, Object, Path))

			' Token: 0x04000042 RID: 66
			Public Shared <>p__Sitea As CallSite(Of Func(Of CallSite, Object, Object))
		End Class
	End Module
End Namespace
