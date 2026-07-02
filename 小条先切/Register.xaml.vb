Imports System
Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Markup

Namespace 小条先切
	' Token: 0x02000007 RID: 7
	Public Partial Class Register
		Inherits Window

		' Token: 0x06000071 RID: 113 RVA: 0x000025F9 File Offset: 0x000007F9
		Public Sub New(cnnStr As String)
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
			Me.InitializeComponent()
			Me.cnnStr = cnnStr
		End Sub

		' Token: 0x06000072 RID: 114 RVA: 0x00002613 File Offset: 0x00000813
		Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
			MyBase.Close()
		End Sub

		' Token: 0x06000073 RID: 115 RVA: 0x0000546C File Offset: 0x0000366C
		Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
			Dim text As String = Me.txtHardInfo.Text.Trim()
			If String.IsNullOrEmpty(text) Then
				MessageBox.Show("请输入机器码,然后注册...", "提示")
				Return
			End If
			If HardInfo.Register(Me.cnnStr, text, "lishuoyang!2", "XKCDMOPTIMIZE") Then
				MessageBox.Show("插件已成功注册!")
				MyBase.Close()
				Return
			End If
			MessageBox.Show("插件注册失败!")
		End Sub

		' Token: 0x0400002D RID: 45
		Private cnnStr As String
	End Class
End Namespace
