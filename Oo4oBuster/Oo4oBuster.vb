Imports Oracle.ManagedDataAccess.Client

Public Class BusterDb
	Implements IDisposable

	Private conn As OracleConnection

	Private Sub New(conn As OracleConnection)
		Me.conn = conn
	End Sub

	''' <summary>
	''' データベース接続を開きます
	''' </summary>
	''' <param name="dbname"></param>
	''' <param name="userid"></param>
	''' <param name="password"></param>
	Public Shared Function OpenDatabase(dbname As String, userid As String, password As String) As BusterDb
		Dim builder = New OracleConnectionStringBuilder() With {
			.UserID = userid,
			.Password = password
			}
		If String.IsNullOrEmpty(dbname) Then
			' DB名が空のときはローカルDBにTNS接続する
			builder.DataSource = "" &
				"(DESCRIPTION = " &
					"(ADDRESS = " &
						"(PROTOCOL = TCP)" &
						"(HOST     = 127.0.0.1)" &
						"(PORT     = 1521)" &
					")" &
				")"
		Else
			builder.DataSource = dbname
		End If
		Dim connectionString = builder.ToString()
		Dim conn = New OracleConnection(connectionString)
		conn.Open()
		Return New BusterDb(conn)
	End Function

	''' <summary>
	''' データベース接続を閉じます。
	''' </summary>
	Public Sub Close()
		If conn IsNot Nothing Then
			conn.Close()
		End If
	End Sub

#Region "IDisposable Support"
	Private disposedValue As Boolean ' 重複する呼び出しを検出するには

	' IDisposable
	Protected Overridable Sub Dispose(disposing As Boolean)
		If Not Me.disposedValue Then
			If disposing Then
				Close()
			End If
		End If
		Me.disposedValue = True
	End Sub

	' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
	Public Sub Dispose() Implements IDisposable.Dispose
		' このコードを変更しないでください。クリーンアップ コードを上の Dispose(disposing As Boolean) に記述します。
		Dispose(True)
		GC.SuppressFinalize(Me)
	End Sub
#End Region

	''' <summary>
	''' 問い合わせSQLをもとにダイナセットを作成します。
	''' </summary>
	Public Function CreateDynaset(sql As String) As Object
		Using cmd As New OracleCommand(sql, conn),
			adapter As New OracleDataAdapter(cmd)

			Dim dt As New DataTable()
			adapter.Fill(dt)

			Return New BusterDynaset(dt)
		End Using
	End Function

	''' <summary>
	''' Dynaset動作を内部クラスでエミュレートし、使用側にObject型として動的アクセスさせます。
	''' </summary>
	Private Class BusterDynaset

		' 内部データストアにはDataTableを使う
		Private store As DataTable
		Sub New(store As DataTable)
			Me.store = store
			If 0 < RecordCount Then
				currentRow = 0
			End If
		End Sub

		' ダイナセットの現在行を保持しておく
		Private currentRow = -1

		Public ReadOnly Property RecordCount
			Get
				Return store.DefaultView.Count
			End Get
		End Property

		Public ReadOnly Property BOF
			Get
				Return currentRow < 0
			End Get
		End Property

		Public ReadOnly Property EOF
			Get
				Return RecordCount <= currentRow
			End Get
		End Property

		' OraFieldへのアクセスをインデクサでエミュレートする
		Default Public ReadOnly Property Item(fieldName As String) As Object
			Get
				Dim row = store.DefaultView(currentRow)
				Return New BusterField(row(fieldName))
			End Get
		End Property

		Default Public ReadOnly Property Item(fieldNo As Integer) As Object
			Get
				Dim row = store.DefaultView(currentRow)
				Return New BusterField(row(fieldNo))
			End Get
		End Property

		' レコード位置の移動は現在行変更で対処する
		Public Sub MoveFirst()
			currentRow = 0
		End Sub

		Public Sub MoveNext()
			currentRow += 1
		End Sub

		Public Sub MovePrevious()
			currentRow -= 1
		End Sub

		Public Sub MoveLast()
			currentRow = RecordCount - 1
		End Sub

		' Find系の絞り込みはDataViewで代替する
		' !!! oo4oでは col1 || col2 = "xxyy"のようなOracle形式の演算子が使えるが、
		' !!! DataViewでは使えないので、 col1 + col2 = "xxyy"のように
		' !!! DataColumns.Expressionの構文で指定する必要がある。
		' !!! https://msdn.microsoft.com/ja-jp/library/system.data.datacolumn.expression.aspx
		Public Sub FindFirst(cond As String)
			store.DefaultView.RowFilter = cond
			If NoMatch Then
				currentRow = -1
			Else
				currentRow = 0
			End If
		End Sub

		Public Sub FindLast(cond As String)
			store.DefaultView.RowFilter = cond
			If NoMatch Then
				currentRow = -1
			Else
				currentRow = store.DefaultView.Count - 1
			End If
		End Sub

		Public ReadOnly Property NoMatch As Boolean
			Get
				Return store.DefaultView.Count = 0
			End Get
		End Property

		Public Sub Close()
			store.Dispose()
		End Sub

		''' <summary>
		''' OraFieldをエミュレートします。
		''' </summary>
		Private Class BusterField
			Private _value As Object
			Sub New(value As Object)
				_value = value
			End Sub
			Public ReadOnly Property Value As Object
				Get
					Return _value
				End Get
			End Property
		End Class

	End Class

End Class
