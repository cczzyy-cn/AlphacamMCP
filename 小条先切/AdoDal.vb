Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Linq
Imports System.Windows

Namespace 小条先切
	' Token: 0x0200000B RID: 11
	Public Module AdoDal
		' Token: 0x06000085 RID: 133 RVA: 0x00005678 File Offset: 0x00003878
		Public Function TestConnectToCDM(cnnStr As String) As Boolean
			Dim flag As Boolean
			Try
				Using oleDbConnection As OleDbConnection = New OleDbConnection(cnnStr)
					oleDbConnection.Open()
				End Using
				flag = True
			Catch ex As Exception
				flag = False
			End Try
			Return flag
		End Function

		' Token: 0x06000086 RID: 134 RVA: 0x000056C8 File Offset: 0x000038C8
		Public Function Find(cnnStr As String, tableName As String, fields As String()) As DataTable
			Dim list As List(Of AdoWhere) = New List(Of AdoWhere)()
			Return AdoDal.Find(cnnStr, tableName, fields, list.ToArray())
		End Function

		' Token: 0x06000087 RID: 135 RVA: 0x000026AF File Offset: 0x000008AF
		Public Function Find(cnnStr As String, tableName As String, fields As String(), ParamArray where As AdoWhere()) As DataTable
			Return AdoDal.Find(cnnStr, tableName, fields, where, True, Nothing)
		End Function

		' Token: 0x06000088 RID: 136 RVA: 0x000056EC File Offset: 0x000038EC
		Public Function Find(cnnStr As String, tableName As String, fields As String(), isASC As Boolean, sortField As String) As DataTable
			Dim list As List(Of AdoWhere) = New List(Of AdoWhere)()
			Return AdoDal.Find(cnnStr, tableName, fields, list.ToArray(), isASC, sortField)
		End Function

		' Token: 0x06000089 RID: 137 RVA: 0x00005710 File Offset: 0x00003910
		Public Function Find(cnnStr As String, tableName As String, fields As String(), where As AdoWhere(), isASC As Boolean, sortField As String) As DataTable
			Dim dataSet As DataSet = New DataSet()
			If fields Is Nothing OrElse fields.Length <= 0 Then
				Return Nothing
			End If
			Dim text As String = "select "
			For i As Integer = 0 To fields.Length - 1
				text = text + "`" + fields(i) + "`,"
			Next
			text = text.Trim(New Char() { ","c })
			text = text + " from `" + tableName + "` "
			Dim list As List(Of OleDbParameter) = New List(Of OleDbParameter)()
			If where IsNot Nothing AndAlso where.Length > 0 Then
				text += "where"
				For j As Integer = 0 To where.Length - 1
					Dim text2 As String = text
					text = String.Concat(New String() { text2, " `", where(j).Field, "`=@", where(j).Field, " and" })
					Dim oleDbParameter As OleDbParameter = New OleDbParameter("@" + where(j).Field, where(j).Value)
					list.Add(oleDbParameter)
				Next
				text = text.TrimEnd("and".ToArray())
			End If
			If Not String.IsNullOrEmpty(sortField) Then
				text = text + "order by `" + sortField + "` "
				If isASC Then
					text += "ASC"
				Else
					text += "DESC"
				End If
			End If
			Try
				Using oleDbConnection As OleDbConnection = New OleDbConnection(cnnStr)
					oleDbConnection.Open()
					Using oleDbCommand As OleDbCommand = oleDbConnection.CreateCommand()
						oleDbCommand.CommandText = text
						oleDbCommand.Parameters.AddRange(list.ToArray())
						Using oleDbDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(oleDbCommand)
							oleDbDataAdapter.Fill(dataSet)
						End Using
					End Using
				End Using
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try
			If dataSet.Tables IsNot Nothing AndAlso dataSet.Tables(0).Rows.Count > 0 Then
				Return dataSet.Tables(0)
			End If
			Return Nothing
		End Function

		' Token: 0x0600008A RID: 138 RVA: 0x00005968 File Offset: 0x00003B68
		Public Function Add(cnnStr As String, tableName As String, ParamArray fields As AdoField()) As Boolean
			If fields IsNot Nothing AndAlso fields.Length > 0 Then
				Dim list As List(Of OleDbParameter) = New List(Of OleDbParameter)()
				Dim text As String = "insert into `" + tableName + "` ("
				Dim text2 As String = ") values("
				For i As Integer = 0 To fields.Length - 1
					text = text + "`" + fields(i).Field + "`,"
					text2 = text2 + "@" + fields(i).Field + ","
					Dim oleDbParameter As OleDbParameter = New OleDbParameter("@" + fields(i).Field, fields(i).Value)
					list.Add(oleDbParameter)
				Next
				text = text.TrimEnd(New Char() { ","c })
				text2 = text2.TrimEnd(New Char() { ","c })
				text2 += ")"
				text += text2
				Dim num As Integer = 0
				Try
					Using oleDbConnection As OleDbConnection = New OleDbConnection(cnnStr)
						oleDbConnection.Open()
						Using oleDbCommand As OleDbCommand = oleDbConnection.CreateCommand()
							oleDbCommand.CommandText = text
							oleDbCommand.Parameters.AddRange(list.ToArray())
							num = oleDbCommand.ExecuteNonQuery()
						End Using
					End Using
				Catch ex As Exception
					MessageBox.Show(ex.Message)
				End Try
				Return num = 1
			End If
			Return False
		End Function

		' Token: 0x0600008B RID: 139 RVA: 0x000026BC File Offset: 0x000008BC
		Public Function FromDbValue(value As Object) As Object
			If value Is DBNull.Value Then
				Return Nothing
			End If
			Return value
		End Function

		' Token: 0x0600008C RID: 140 RVA: 0x00005AEC File Offset: 0x00003CEC
		Public Sub Adds(cnnStr As String, tableName As String, fields As List(Of List(Of AdoField)))
			If fields IsNot Nothing AndAlso fields.Count > 0 Then
				Try
					Using oleDbConnection As OleDbConnection = New OleDbConnection(cnnStr)
						oleDbConnection.Open()
						Using oleDbCommand As OleDbCommand = oleDbConnection.CreateCommand()
							For i As Integer = 0 To fields.Count - 1
								oleDbCommand.CommandText = Nothing
								oleDbCommand.Parameters.Clear()
								Dim array As AdoField() = fields(i).ToArray()
								Dim list As List(Of OleDbParameter) = New List(Of OleDbParameter)()
								Dim text As String = "insert into `" + tableName + "` ("
								Dim text2 As String = ") values("
								For j As Integer = 0 To array.Length - 1
									text = text + "`" + array(j).Field + "`,"
									text2 = text2 + "@" + array(j).Field + ","
									Dim oleDbParameter As OleDbParameter = New OleDbParameter("@" + array(j).Field, array(j).Value)
									list.Add(oleDbParameter)
								Next
								text = text.TrimEnd(New Char() { ","c })
								text2 = text2.TrimEnd(New Char() { ","c })
								text2 += ")"
								text += text2
								oleDbCommand.CommandText = text
								oleDbCommand.Parameters.AddRange(list.ToArray())
								oleDbCommand.ExecuteNonQuery()
							Next
						End Using
					End Using
				Catch ex As Exception
					MessageBox.Show(ex.Message)
				End Try
				Return
			End If
		End Sub

		' Token: 0x0600008D RID: 141 RVA: 0x00005CD0 File Offset: 0x00003ED0
		Public Function HasTable(cnnStr As String, tableName As String) As Boolean
			Dim dataTable As DataTable = Nothing
			Dim flag As Boolean = False
			Try
				Using oleDbConnection As OleDbConnection = New OleDbConnection(cnnStr)
					oleDbConnection.Open()
					dataTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() { Nothing, Nothing, Nothing, "TABLE" })
				End Using
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try
			If dataTable IsNot Nothing AndAlso dataTable.Rows.Count > 0 Then
				For Each obj As Object In dataTable.Rows
					Dim dataRow As DataRow = CType(obj, DataRow)
					If String.Compare(dataRow("TABLE_NAME").ToString().Trim(), tableName.Trim(), True) = 0 Then
						flag = True
						Exit For
					End If
				Next
			End If
			Return flag
		End Function

		' Token: 0x0600008E RID: 142 RVA: 0x00005DCC File Offset: 0x00003FCC
		Public Function ExcuteSql(cnnStr As String, sql As String) As Boolean
			Dim flag As Boolean = False
			Try
				Using oleDbConnection As OleDbConnection = New OleDbConnection(cnnStr)
					oleDbConnection.Open()
					Using oleDbCommand As OleDbCommand = oleDbConnection.CreateCommand()
						oleDbCommand.CommandText = sql
						oleDbCommand.ExecuteNonQuery()
						flag = True
					End Using
				End Using
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try
			Return flag
		End Function

		' Token: 0x0600008F RID: 143 RVA: 0x00005E54 File Offset: 0x00004054
		Public Sub CreateTableAD_DOOR_XKSET(cnnStr As String)
			Dim text As String = "create table AD_DOOR_XKSET (ID AUTOINCREMENT(1,1) primary key, Project  char(50), SubProject char(50))"
			AdoDal.ExcuteSql(cnnStr, text)
			text = "alter table AD_DOOR_XKSET add column Value1 char(255)"
			AdoDal.ExcuteSql(cnnStr, text)
			text = "alter table AD_DOOR_XKSET add column Value2 char(255)"
			AdoDal.ExcuteSql(cnnStr, text)
			text = "alter table AD_DOOR_XKSET add column Value3 char(255)"
			AdoDal.ExcuteSql(cnnStr, text)
			text = "alter table AD_DOOR_XKSET add column Value4 char(255)"
			AdoDal.ExcuteSql(cnnStr, text)
			text = "alter table AD_DOOR_XKSET add column Value5 char(255)"
			AdoDal.ExcuteSql(cnnStr, text)
		End Sub

		' Token: 0x06000090 RID: 144 RVA: 0x00005EB8 File Offset: 0x000040B8
		Public Function Save(cnnStr As String, tableName As String, fields As AdoField(), ParamArray where As AdoWhere()) As Boolean
			Dim num As Integer = 0
			Dim text As String = "update " + tableName + " set "
			Dim list As List(Of OleDbParameter) = New List(Of OleDbParameter)()
			For i As Integer = 0 To fields.Length - 1
				Dim text2 As String = text
				text = String.Concat(New String() { text2, "`", fields(i).Field, "` = @", fields(i).Field, "," })
				Dim oleDbParameter As OleDbParameter = New OleDbParameter("@" + fields(i).Field, fields(i).Value)
				list.Add(oleDbParameter)
			Next
			text = text.TrimEnd(New Char() { ","c }) + " where"
			For j As Integer = 0 To where.Length - 1
				Dim text3 As String = text
				text = String.Concat(New String() { text3, " `", where(j).Field, "`=@", where(j).Field, " and" })
				Dim oleDbParameter2 As OleDbParameter = New OleDbParameter("@" + where(j).Field, where(j).Value)
				list.Add(oleDbParameter2)
			Next
			text = text.TrimEnd("and".ToArray())
			Using oleDbConnection As OleDbConnection = New OleDbConnection(cnnStr)
				oleDbConnection.Open()
				Using oleDbCommand As OleDbCommand = oleDbConnection.CreateCommand()
					oleDbCommand.CommandText = text
					oleDbCommand.Parameters.AddRange(list.ToArray())
					num = oleDbCommand.ExecuteNonQuery()
				End Using
			End Using
			Return num = 1
		End Function
	End Module
End Namespace
