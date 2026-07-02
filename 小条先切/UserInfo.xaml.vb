Imports System
Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Markup

Namespace 小条先切
	' Token: 0x02000008 RID: 8
	Public Partial Class UserInfo
		Inherits Window

		' Token: 0x06000076 RID: 118 RVA: 0x0000261B File Offset: 0x0000081B
		Public Sub New(user As User)
			Class2.C5hn1h9zVegKm()
			MyBase..ctor()
			Me.info = user
			Me.InitializeComponent()
		End Sub

		' Token: 0x06000077 RID: 119 RVA: 0x00002613 File Offset: 0x00000813
		Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
			MyBase.Close()
		End Sub

		' Token: 0x06000078 RID: 120 RVA: 0x00005570 File Offset: 0x00003770
		Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
			Me.txtAuthorGuy.Text = Me.info.AuthorGuy
			Dim dateTime As DateTime = New DateTime(Me.info.DueDate)
			Me.txtDueDay.Text = dateTime.ToLongDateString() + dateTime.ToLongTimeString()
		End Sub

		' Token: 0x04000030 RID: 48
		Private info As User
	End Class
End Namespace
