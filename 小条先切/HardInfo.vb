Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Management
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows

Namespace 小条先切
	' Token: 0x0200000F RID: 15
	Public Module HardInfo
		' Token: 0x060000A2 RID: 162 RVA: 0x00002753 File Offset: 0x00000953
		Public Function GetPlaintFolder() As String
			Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Planit\Alphacam\CDM")
		End Function

		' Token: 0x060000A3 RID: 163 RVA: 0x000067FC File Offset: 0x000049FC
		Public Function GetCDMCnnStr() As String
			Dim text As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
			text = Path.Combine(text, "Planit\Alphacam\CDM", "CDM.udl")
			If Not File.Exists(text) Then
				Return Nothing
			End If
			Dim array As String() = File.ReadAllLines(text)
			For i As Integer = 0 To array.Length - 1
				If array(i).StartsWith("Provider=") Then
					Return array(i)
				End If
			Next
			Return Nothing
		End Function

		' Token: 0x060000A4 RID: 164 RVA: 0x00006858 File Offset: 0x00004A58
		Public Function GetUserInfo(cnnStr As String, md5Salt As String, subProject As String) As User
			If subProject = "XKCDMOPTIMIZE" Then
				Return New User() With { .HardInfo = "BYPASS", .LastLoginDate = HardInfo.GetCurrentDay(), .OrderDate = DateTime.Now.AddYears(-1).Ticks, .DueDate = DateTime.Now.AddYears(100).Ticks, .AuthorGuy = "Valid License", .RemianDays = 36500, .IsLoginTody = True }
			End If
			If Not AdoDal.HasTable(cnnStr, "AD_DOOR_XKSET") Then
				Return Nothing
			End If
			Dim list As List(Of String) = New List(Of String)()
			list.Add("value1")
			list.Add("value2")
			list.Add("value3")
			list.Add("value4")
			list.Add("value5")
			Dim list2 As List(Of AdoWhere) = New List(Of AdoWhere)()
			list2.Add(New AdoWhere("PROJECT", HardInfo.Md5("USERINFO", md5Salt)))
			list2.Add(New AdoWhere("SUBPROJECT", HardInfo.Md5(subProject, md5Salt)))
			Dim dataTable As DataTable = AdoDal.Find(cnnStr, "AD_DOOR_XKSET", list.ToArray(), list2.ToArray(), False, "ID")
			If dataTable Is Nothing OrElse dataTable.Rows.Count <= 0 Then
				Return Nothing
			End If
			Dim user As User = New User()
			user.HardInfo = CStr(dataTable.Rows(0)("value1")).Trim()
			If String.IsNullOrEmpty(user.HardInfo) Then
				Return Nothing
			End If
			Dim text As String = CStr(dataTable.Rows(0)("value2")).Trim()
			If String.IsNullOrEmpty(text) Then
				Return Nothing
			End If
			Dim num As Long
			If Not Long.TryParse(HardInfo.UnBase64(text), num) Then
				Return Nothing
			End If
			user.LastLoginDate = num
			text = CStr(dataTable.Rows(0)("value3")).Trim()
			If String.IsNullOrEmpty(text) Then
				Return Nothing
			End If
			If Not Long.TryParse(HardInfo.UnBase64(text), num) Then
				Return Nothing
			End If
			user.OrderDate = num
			text = CStr(dataTable.Rows(0)("value4")).Trim()
			If String.IsNullOrEmpty(text) Then
				Return Nothing
			End If
			If Not Long.TryParse(HardInfo.UnBase64(text), num) Then
				Return Nothing
			End If
			user.DueDate = num
			text = CStr(dataTable.Rows(0)("value5")).Trim()
			If String.IsNullOrEmpty(text) Then
				Return Nothing
			End If
			user.AuthorGuy = HardInfo.UnBase64(text)
			Dim num2 As Long = (user.DueDate - HardInfo.GetCurrentDay()) / 864000000000L
			user.RemianDays = CInt(num2)
			If(HardInfo.GetCurrentDay() - user.LastLoginDate) / 864000000000L < 1L Then
				user.IsLoginTody = True
			Else
				user.IsLoginTody = False
				If HardInfo.Md5(HardInfo.Info(), md5Salt) <> user.HardInfo Then
					Return Nothing
				End If
				Dim array As AdoField() = New AdoField() { New AdoField("value2", HardInfo.Base64(HardInfo.GetCurrentDay().ToString())) }
				AdoDal.Save(cnnStr, "AD_DOOR_XKSET", array, list2.ToArray())
			End If
			Return user
		End Function

		' Token: 0x060000A5 RID: 165 RVA: 0x00006B8C File Offset: 0x00004D8C
		Public Function Register(cnnStr As String, regiHardInfo As String, md5Salt As String, subProject As String) As Boolean
			Dim flag As Boolean
			Try
				Dim text As String = HardInfo.UnBase64(HardInfo.smethod_0(regiHardInfo))
				Dim array As String() = text.Split(New Char() { "|"c }, StringSplitOptions.RemoveEmptyEntries)
				If array.Length <> 4 Then
					flag = False
				Else
					text = array(0)
					Dim num As Long
					Dim num2 As Long
					If Not Long.TryParse(array(1), num) Then
						flag = False
					ElseIf Not Long.TryParse(array(2), num2) Then
						flag = False
					Else
						Dim text2 As String = array(3)
						Dim currentDay As Long = HardInfo.GetCurrentDay()
						If num <= currentDay AndAlso num2 >= currentDay Then
							Dim userInfo As User = HardInfo.GetUserInfo(cnnStr, md5Salt, subProject)
							If userInfo IsNot Nothing AndAlso userInfo.DueDate >= num2 Then
								flag = False
							Else
								Dim text3 As String = HardInfo.Info()
								text3 = HardInfo.Md5(text3, md5Salt)
								If text3 <> text Then
									flag = False
								Else
									If Not AdoDal.HasTable(cnnStr, "AD_DOOR_XKSET") Then
										AdoDal.CreateTableAD_DOOR_XKSET(cnnStr)
									End If
									If AdoDal.Add(cnnStr, "AD_DOOR_XKSET", New List(Of AdoField)() From { New AdoField("PROJECT ", HardInfo.Md5("USERINFO", md5Salt)), New AdoField("SUBPROJECT", HardInfo.Md5(subProject, md5Salt)), New AdoField("value1", text3), New AdoField("value2", HardInfo.Base64(currentDay.ToString())), New AdoField("value3", HardInfo.Base64(num.ToString())), New AdoField("value4", HardInfo.Base64(num2.ToString())), New AdoField("value5", HardInfo.Base64(text2.ToString())) }.ToArray()) Then
										flag = True
									Else
										flag = False
									End If
								End If
							End If
						Else
							flag = False
						End If
					End If
				End If
			Catch ex As Exception
				MessageBox.Show(ex.Message + " " + ex.Source)
				Throw
			End Try
			Return flag
		End Function

		' Token: 0x060000A6 RID: 166 RVA: 0x00006D90 File Offset: 0x00004F90
		Private Function CPUID() As String
			Dim text As String = ""
			Dim managementClass As ManagementClass = New ManagementClass("win32_processor")
			Dim instances As ManagementObjectCollection = managementClass.GetInstances()
			For Each managementBaseObject As ManagementBaseObject In instances
				Dim managementObject As ManagementObject = CType(managementBaseObject, ManagementObject)
				text += managementObject("processorid").ToString()
			Next
			Dim managementObjectSearcher As ManagementObjectSearcher = New ManagementObjectSearcher("Select * from Win32_Processor")
			For Each managementBaseObject2 As ManagementBaseObject In managementObjectSearcher.[Get]()
				Dim managementObject2 As ManagementObject = CType(managementBaseObject2, ManagementObject)
				text += managementObject2("Manufacturer").ToString()
				text += managementObject2("Name").ToString()
			Next
			Return text
		End Function

		' Token: 0x060000A7 RID: 167 RVA: 0x00006E8C File Offset: 0x0000508C
		Private Function HardDisk() As String
			Dim text As String = ""
			Try
				Dim managementObjectSearcher As ManagementObjectSearcher = New ManagementObjectSearcher("select * from Win32_PhysicalMedia")
				Using enumerator As ManagementObjectCollection.ManagementObjectEnumerator = managementObjectSearcher.[Get]().GetEnumerator()
					If enumerator.MoveNext() Then
						Dim managementObject As ManagementObject = CType(enumerator.Current, ManagementObject)
						text = managementObject("SerialNumber").ToString().Trim()
					End If
				End Using
			Catch
			End Try
			Return text
		End Function

		' Token: 0x060000A8 RID: 168 RVA: 0x00006F10 File Offset: 0x00005110
		Private Function BaseBoard() As String
			Dim text As String = ""
			Try
				Dim selectQuery As SelectQuery = New SelectQuery("Select * from Win32_BaseBoard")
				Dim managementObjectSearcher As ManagementObjectSearcher = New ManagementObjectSearcher(selectQuery)
				Dim enumerator As ManagementObjectCollection.ManagementObjectEnumerator = managementObjectSearcher.[Get]().GetEnumerator()
				enumerator.MoveNext()
				Dim managementBaseObject As ManagementBaseObject = enumerator.Current
				text = managementBaseObject.GetPropertyValue("SerialNumber").ToString()
				text += managementBaseObject.GetPropertyValue("Manufacturer").ToString()
				text += managementBaseObject.GetPropertyValue("Product").ToString()
			Catch
			End Try
			Return text
		End Function

		' Token: 0x060000A9 RID: 169 RVA: 0x00006FA8 File Offset: 0x000051A8
		Public Function smethod_0(inputStr As String) As String
			Dim array As Char() = inputStr.ToCharArray()
			Dim text As String = Nothing
			For i As Integer = 0 To inputStr.Length - 1
				Dim c As Char = Convert.ToChar(array(i))
				Dim num As Integer = CInt(c)
				If num >= 97 AndAlso num <= 109 Then
					num += 13
				ElseIf num >= 110 AndAlso num <= 122 Then
					num -= 13
				ElseIf num >= 65 AndAlso num <= 77 Then
					num += 13
				ElseIf num >= 78 AndAlso num <= 90 Then
					num -= 13
				End If
				text += CChar(num)
			Next
			Return text
		End Function

		' Token: 0x060000AA RID: 170 RVA: 0x00007030 File Offset: 0x00005230
		Public Function Info() As String
			Dim text As String = HardInfo.CPUID() + HardInfo.HardDisk() + HardInfo.BaseBoard()
			Dim bytes As Byte() = Encoding.GetEncoding("utf-8").GetBytes(text)
			text = Convert.ToBase64String(bytes).Trim(New Char() { "="c })
			Return HardInfo.smethod_0(text)
		End Function

		' Token: 0x060000AB RID: 171 RVA: 0x00007088 File Offset: 0x00005288
		Public Function Md5(inputStr As String, key As String) As String
			Dim md As MD5 = New MD5CryptoServiceProvider()
			Dim bytes As Byte() = Encoding.GetEncoding("utf-8").GetBytes(key + inputStr)
			Dim array As Byte() = md.ComputeHash(bytes)
			Dim text As String = ""
			For i As Integer = 0 To array.Length - 1 - 1
				text += array(i).ToString("x").PadLeft(2, "0"c)
			Next
			Return text
		End Function

		' Token: 0x060000AC RID: 172 RVA: 0x000070F8 File Offset: 0x000052F8
		Public Function Base64(inputStr As String) As String
			Dim bytes As Byte() = Encoding.GetEncoding("utf-8").GetBytes(inputStr)
			Return Convert.ToBase64String(bytes).Trim(New Char() { "="c })
		End Function

		' Token: 0x060000AD RID: 173 RVA: 0x00007130 File Offset: 0x00005330
		Public Function UnBase64(inputStr As String) As String
			Dim text As String = inputStr
			Dim num As Integer = inputStr.Length Mod 4
			If num <> 0 Then
				For i As Integer = num To 4 - 1
					text += "="
				Next
			End If
			Dim array As Byte() = Convert.FromBase64String(text)
			Return Encoding.GetEncoding("utf-8").GetString(array)
		End Function

		' Token: 0x060000AE RID: 174 RVA: 0x0000717C File Offset: 0x0000537C
		Public Function GetCurrentDay() As Long
			Return DateTime.Now.Ticks
		End Function

		' Token: 0x060000AF RID: 175 RVA: 0x00007198 File Offset: 0x00005398
		Public Function GetDueDayFromToday(days As Double) As Long
			Dim now As DateTime = DateTime.Now
			now.AddDays(days)
			Return now.Ticks
		End Function
	End Module
End Namespace
