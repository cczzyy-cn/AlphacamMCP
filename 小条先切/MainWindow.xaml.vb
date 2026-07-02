Imports System
Imports System.CodeDom.Compiler
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Markup
Imports AlphaCAMRouter

Namespace 小条先切
	' Token: 0x02000012 RID: 18
	Public Partial Class MainWindow
		Inherits Window

		' Token: 0x060000B7 RID: 183 RVA: 0x00002794 File Offset: 0x00000994
		Public Sub New()
			Class2.C5hn1h9zVegKm()
			Me.toolNames = New List(Of String)()
			MyBase..ctor()
			Me.InitializeComponent()
		End Sub

		' Token: 0x060000B8 RID: 184 RVA: 0x0000724C File Offset: 0x0000544C
		Private Sub muAbout_Click(sender As Object, e As RoutedEventArgs)
			Dim about As About = New About()
			about.ShowDialog()
		End Sub

		' Token: 0x060000B9 RID: 185 RVA: 0x00007268 File Offset: 0x00005468
		Private Sub muRegister_Click(sender As Object, e As RoutedEventArgs)
			Dim register As Register = New Register(Me.cnnStr)
			register.ShowDialog()
		End Sub

		' Token: 0x060000BA RID: 186 RVA: 0x00007288 File Offset: 0x00005488
		Private Sub muUserInfo_Click(sender As Object, e As RoutedEventArgs)
			Dim userInfo As UserInfo = New UserInfo(Me.userInfo)
			userInfo.ShowDialog()
		End Sub

		' Token: 0x060000BB RID: 187 RVA: 0x000072A8 File Offset: 0x000054A8
		Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
			Dim flag As Boolean = True
			Me.cnnStr = HardInfo.GetCDMCnnStr()
			If String.IsNullOrEmpty(Me.cnnStr) OrElse Not AdoDal.TestConnectToCDM(Me.cnnStr) Then
				MessageBox.Show("不能连接到您的CDM数据库..." + Environment.NewLine + Environment.NewLine + "请确认您的数据库配置", "出错提示！")
				MyBase.Close()
			End If
			Me.userInfo = HardInfo.GetUserInfo(Me.cnnStr, "lishuoyang!2", "XKCDMOPTIMIZE")
			If Me.userInfo Is Nothing Then
				flag = False
				Me.muUserInfo.IsEnabled = False
				Me.txtblkPrompt.Text = "您的插件没有注册,请在帮助里注册该插件..."
				MessageBox.Show("您的插件没有注册,请在帮助里注册该插件...", "提示")
			End If
			Dim currentDay As Long = HardInfo.GetCurrentDay()
			If flag AndAlso (currentDay < Me.userInfo.LastLoginDate OrElse currentDay > Me.userInfo.DueDate) Then
				flag = False
				Dim text As String = Environment.NewLine + Environment.NewLine + Me.userInfo.AuthorGuy
				Me.txtblkPrompt.Text = "您插件的注册信息已失效..."
				MessageBox.Show("您插件的注册信息已失效,请联系您的授权方：" + text, "提示")
			End If
			If flag AndAlso Me.userInfo.RemianDays < 30 Then
				Dim text2 As String = Environment.NewLine + Environment.NewLine + Me.userInfo.AuthorGuy
				MessageBox.Show(String.Concat(New Object() { "您插件将于 ", Me.userInfo.RemianDays, " 天后停止使用 ,请联系您的授权方：", text2 }), "提示")
			End If
			If Not flag Then
				Me.btnOK.IsEnabled = False
				Return
			End If
			Me.Acam = AlphaTool.GetAlphaApp()
			If Me.Acam Is Nothing Then
				Me.txtblkPrompt.Text = "连接Alphacam服务失败!"
				Me.btnOK.IsEnabled = False
				Return
			End If
			If Not AlphaTool.HasNesting(Me.Acam) Then
				Me.txtblkPrompt.Text = "未能在当前图形中发现排版信息!"
				Me.btnOK.IsEnabled = False
				Return
			End If
			Me.txtblkPrompt.Text = "选择切断刀,然后优化,优化中请耐心等待..."
			Me.toolNames = AlphaTool.GetDrwToolName(Me.Acam)
			If Me.toolNames.Count <= 0 Then
				Me.txtblkPrompt.Text = "未能在当前图形中发现刀具信息!"
				Me.btnOK.IsEnabled = False
				Return
			End If
			Me.cmbToolName.ItemsSource = Me.toolNames
			Me.userSet = New [Set]()
			Me.GetCDMUserSet()
			If Me.set_0 IsNot Nothing Then
				Me.userSet.Copy(Me.set_0)
			End If
			Me.grdSet.DataContext = Me.userSet
			Me.chbUniCutTool.DataContext = Me.userSet
			Me.cmbToolName.DataContext = Me.userSet
		End Sub

		' Token: 0x060000BC RID: 188 RVA: 0x00007558 File Offset: 0x00005758
		Private Sub SaveUserSet()
			If Me.set_0 Is Nothing OrElse Not Me.set_0.Equal(Me.userSet) Then
				If Me.set_0 IsNot Nothing Then
					Dim list As List(Of AdoWhere) = New List(Of AdoWhere)()
					list.Add(New AdoWhere("PROJECT", HardInfo.Md5("USERSET", "lishuoyang!2")))
					list.Add(New AdoWhere("SUBPROJECT", HardInfo.Md5("SAMALLFIRSTCUT", "lishuoyang!2")))
					Dim list2 As List(Of AdoField) = New List(Of AdoField)()
					Dim text As String = Me.userSet.ToString()
					text = HardInfo.Base64(text)
					list2.Add(New AdoField("value1", text))
					Dim text2 As String = Me.userSet.BackDist.ToString()
					text2 = HardInfo.Base64(text2)
					list2.Add(New AdoField("value2", text2))
					AdoDal.Save(Me.cnnStr, "AD_DOOR_XKSET", list2.ToArray(), list.ToArray())
					Return
				End If
				Dim list3 As List(Of AdoField) = New List(Of AdoField)()
				list3.Add(New AdoField("PROJECT", HardInfo.Md5("USERSET", "lishuoyang!2")))
				list3.Add(New AdoField("SUBPROJECT", HardInfo.Md5("SAMALLFIRSTCUT", "lishuoyang!2")))
				Dim text3 As String = Me.userSet.ToString()
				text3 = HardInfo.Base64(text3)
				list3.Add(New AdoField("value1", text3))
				Dim text4 As String = Me.userSet.BackDist.ToString()
				text4 = HardInfo.Base64(text4)
				list3.Add(New AdoField("value2", text4))
				AdoDal.Add(Me.cnnStr, "AD_DOOR_XKSET", list3.ToArray())
			End If
		End Sub

		' Token: 0x060000BD RID: 189 RVA: 0x00007704 File Offset: 0x00005904
		Private Sub GetCDMUserSet()
			Dim list As List(Of AdoWhere) = New List(Of AdoWhere)()
			list.Add(New AdoWhere("PROJECT", HardInfo.Md5("USERSET", "lishuoyang!2")))
			list.Add(New AdoWhere("SUBPROJECT", HardInfo.Md5("SAMALLFIRSTCUT", "lishuoyang!2")))
			Dim dataTable As DataTable = AdoDal.Find(Me.cnnStr, "AD_DOOR_XKSET", New String() { "value1", "value2" }, list.ToArray(), False, "ID")
			If dataTable IsNot Nothing AndAlso dataTable.Rows.Count >= 1 Then
				Dim text As String = CStr(dataTable.Rows(0)("value1")).Trim()
				If Not String.IsNullOrEmpty(text) Then
					text = HardInfo.UnBase64(text)
					Dim array As String() = text.Split(New Char() { "|"c }, StringSplitOptions.RemoveEmptyEntries)
					If array.Length = 5 Then
						Me.set_0 = New [Set]()
						Me.set_0.ToolName = array(0)
						Me.set_0.SmallWidth = Double.Parse(array(1))
						If String.Compare(array(2), "True", True) = 0 Then
							Me.set_0.IsSloopingLine = True
						Else
							Me.set_0.IsSloopingLine = False
						End If
						Me.set_0.SloopDist = Double.Parse(array(3))
						If String.Compare(array(4), "True") = 0 Then
							Me.set_0.IsUniqueCutTool = True
						Else
							Me.set_0.IsUniqueCutTool = False
						End If
						If dataTable.Rows(0)("value2") IsNot DBNull.Value Then
							text = CStr(dataTable.Rows(0)("value2")).Trim()
							Me.set_0.BackDist = Double.Parse(HardInfo.UnBase64(text))
							Return
						End If
						Me.set_0.BackDist = 0.0
					End If
				End If
			End If
		End Sub

		' Token: 0x060000BE RID: 190 RVA: 0x000078F4 File Offset: 0x00005AF4
		Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
			If String.IsNullOrEmpty(Me.selectedToolName) Then
				MessageBox.Show("请选择切断刀...", "提示")
				Return
			End If
			Me.SaveUserSet()
			Me.txtblkPrompt.Text = "正在提取排版信息..."
			Try
				Me.drwNests = AlphaTool.GetDrwNesting(Me.Acam)
			Catch ex As Exception
				MessageBox.Show(ex.Message + "e2 11111" + ex.Source)
				Return
			End Try
			Me.txtblkPrompt.Text = "正在提取刀路信息..."
			Me.drwCutToolPaths = AlphaTool.GetToolGeoByToolName(Me.Acam.ActiveDrawing, Me.selectedToolName, False, False)
			If Me.AssignToolPathToNest() <= 0 Then
				Me.txtblkPrompt.Text = "没有发现符合要求的刀路信息!"
				MessageBox.Show("没有发现符合要求的刀路信息!", "提示")
			Else
				If Not File.Exists(Me.drwNests.ElementAt(0).ToolPaths.ElementAt(0).ToolFileName) Then
					Me.txtblkPrompt.Text = "要求的刀具文件不存在！"
					MessageBox.Show("要求的刀具文件不存在！" + Environment.NewLine + Me.drwNests.ElementAt(0).ToolPaths.ElementAt(0).ToolFileName, "提示")
					Return
				End If
				Me.txtblkPrompt.Text = "正在优化,请耐心等待 ..."
				For i As Integer = 0 To Me.drwNests.Count - 1
					Me.drwNests(i).Acam = Me.Acam
					Me.drwNests(i).IsNestH = True
					If Me.drwNests.ElementAt(i).MaxY - Me.drwNests.ElementAt(i).MinY > Me.drwNests.ElementAt(i).MaxX - Me.drwNests.ElementAt(i).MinX Then
						Me.drwNests(i).IsNestH = False
					End If
					Me.drwNests(i).SloopingLine = Me.userSet.IsSloopingLine
					Me.drwNests(i).SmallPanel = Me.userSet.SmallWidth
					Me.drwNests(i).SloopDist = Me.userSet.SloopDist
					Me.drwNests(i).BackDist = Me.userSet.BackDist
					Me.drwNests(i).UniqueCutTool = Me.userSet.IsUniqueCutTool
					Me.drwNests(i).OptimizeCutPaths()
				Next
				Me.Acam.ActiveDrawing.OrderToolPathsInNestedSheets()
				Me.txtblkPrompt.Text = "优化完成!"
				MessageBox.Show("优化完成!")
				Return
			End If
		End Sub

		' Token: 0x060000BF RID: 191 RVA: 0x00007BB0 File Offset: 0x00005DB0
		Private Function AssignToolPathToNest() As Integer
			Dim num As Integer = 0
			Dim i As Integer = 0
			IL_00F6:
			While i < Me.drwCutToolPaths.Count
				Dim num2 As Double = (Me.drwCutToolPaths.ElementAt(i).MinX + Me.drwCutToolPaths.ElementAt(i).MaxX) / 2.0
				Dim num3 As Double = (Me.drwCutToolPaths.ElementAt(i).MinY + Me.drwCutToolPaths.ElementAt(i).MaxY) / 2.0
				For j As Integer = 0 To Me.drwNests.Count - 1
					If Math.Abs(Me.drwCutToolPaths.ElementAt(i).FinalDepth) >= Me.drwNests.ElementAt(j).Thickness * 0.9 AndAlso Me.drwNests.ElementAt(j).TextInside(num2, num3) Then
						Me.drwNests.ElementAt(j).ToolPaths.Add(Me.drwCutToolPaths.ElementAt(i))
						num += 1
						IL_00F2:
						i += 1
						GoTo IL_00F6
					End If
				Next
				GoTo IL_00F2
			End While
			Me.drwCutToolPaths.Clear()
			Return num
		End Function

		' Token: 0x060000C0 RID: 192 RVA: 0x00002613 File Offset: 0x00000813
		Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs)
			MyBase.Close()
		End Sub

		' Token: 0x060000C1 RID: 193 RVA: 0x000027B2 File Offset: 0x000009B2
		Private Sub cmbToolName_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
			If Me.cmbToolName.SelectedItem IsNot Nothing Then
				Me.selectedToolName = CStr(Me.cmbToolName.SelectedItem)
			End If
		End Sub

		' Token: 0x060000C2 RID: 194 RVA: 0x00007CD0 File Offset: 0x00005ED0
		Private Sub CheckBox_Click(sender As Object, e As RoutedEventArgs)
			If Me.chkSlooping.IsChecked = True Then
				Me.txtSloopDist.IsEnabled = True
				Return
			End If
			Me.txtSloopDist.IsEnabled = False
		End Sub

		' Token: 0x0400004A RID: 74
		Private cnnStr As String

		' Token: 0x0400004B RID: 75
		Private userInfo As User

		' Token: 0x0400004C RID: 76
		Private Acam As App

		' Token: 0x0400004D RID: 77
		Private toolNames As List(Of String)

		' Token: 0x0400004E RID: 78
		Private selectedToolName As String

		' Token: 0x0400004F RID: 79
		Private drwNests As List(Of DrwNest)

		' Token: 0x04000050 RID: 80
		Private drwCutToolPaths As List(Of CutToolPath)

		' Token: 0x04000051 RID: 81
		Private userSet As [Set]

		' Token: 0x04000052 RID: 82
		Private set_0 As [Set]
	End Class
End Namespace
