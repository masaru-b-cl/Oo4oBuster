Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Runtime.InteropServices
Imports System.Configuration

<TestClass()>
Public Class Oo4oBusterTest

	Private Shared dbName As String
	Private Shared userId As String
	Private Shared password As String

	<ClassInitialize>
	Public Shared Sub SetUpClass(context As TestContext)
		dbName = ConfigurationManager.AppSettings("dbName")
		userId = ConfigurationManager.AppSettings("userid")
		password = ConfigurationManager.AppSettings("password")
	End Sub

	<TestMethod()>
	Public Sub TestOo4o()
		Const ORADB_DEFAULT = 0
		Const ORADYN_READONLY = 4

		Dim session As Object = Nothing
		Dim db As Object = Nothing
		Dim dynaset As Object = Nothing

		Try
			session = CreateObject("OracleInProcServer.XOraSession")
			db = session.OpenDatabase(dbName, userId & "/" & password, ORADB_DEFAULT)
			dynaset = db.CreateDynaset("" &
				" select 'sho' as name, 35 as age from dual" &
				" union all" &
				" select null as name, null as age from dual" &
				"", ORADYN_READONLY)

			Assert.AreEqual(2, dynaset.RecordCount)

			Assert.IsFalse(dynaset.EOF)

			Dim name As Object = Nothing
			Dim age As Object = Nothing

			Try
				name = dynaset("name")
				Assert.AreEqual("sho", name.Value)
			Finally
				Marshal.ReleaseComObject(name)
			End Try

			Try
				age = dynaset("age")
				Assert.IsTrue(age.Value = 35)
			Finally
				Marshal.ReleaseComObject(age)
			End Try

			dynaset.MoveNext()

			Assert.IsFalse(dynaset.EOF)

			Try
				name = dynaset("name")
				Assert.IsInstanceOfType(name.Value, GetType(DBNull))
			Finally
				Marshal.ReleaseComObject(name)
			End Try

			Try
				age = dynaset("age")
				Assert.IsInstanceOfType(age.Value, GetType(DBNull))
			Finally
				Marshal.ReleaseComObject(age)
			End Try

			dynaset.MoveNext()

			Assert.IsTrue(dynaset.EOF)
		Finally
			If dynaset IsNot Nothing Then
				dynaset.Close()
				Marshal.ReleaseComObject(dynaset)
			End If
			If db IsNot Nothing Then
				db.Close()
				Marshal.ReleaseComObject(db)
			End If
			If session IsNot Nothing Then
				Marshal.ReleaseComObject(session)
			End If
		End Try

	End Sub


	<TestMethod()>
	Public Sub TestBuster()
		Using db = BusterDb.OpenDatabase(dbName, userId, password)
			Dim dynaset = db.CreateDynaset("" &
				" select 'sho' as name, 35 as age from dual" &
				" union all" &
				" select null as name, null as age from dual" &
				"")

			Assert.IsFalse(dynaset.EOF)

			Assert.AreEqual(2, dynaset.RecordCount)

			Assert.AreEqual("sho", dynaset("name").Value)
			Assert.IsTrue(dynaset("age").Value = 35)

			dynaset.MoveNext()

			Assert.IsFalse(dynaset.EOF)

			Assert.IsInstanceOfType(dynaset("name").Value, GetType(DBNull))
			Assert.IsInstanceOfType(dynaset("age").Value, GetType(DBNull))

			dynaset.MoveNext()

			Assert.IsTrue(dynaset.EOF)

			db.Close()
		End Using
	End Sub


End Class