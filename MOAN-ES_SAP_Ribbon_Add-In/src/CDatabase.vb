Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SQLite
Imports System.Diagnostics
Imports System.IO

Public Class CDatabase : Implements IDisposable
    ' Database wrapper for SQLite3 (ADO)
    ' @see more CRUD http://www.dreamincode.net/forums/topic/157830-using-sqlite-with-c%23/
    ' ADO.DB connectors for System.Data.SQLite: https://system.data.sqlite.org/index.html/doc/trunk/www/downloads.wiki

    Private dbConnection As String
    Private disposedValue As Boolean ' To detect redundant calls

    ' Default constructor
    ' Return void. 
    Public Sub New()
        Dim dbPath As String = My.Settings.AppDbPath
        dbConnection = "Data Source=" & dbPath
    End Sub

    ' Overloading constructor - set another file name.
    ' Param String dbFullPath, the path to a database. 
    ' Return void. 
    Public Sub New(ByVal dbFullPath As String)
        dbConnection = "Data Source=" & dbFullPath
    End Sub

    ' Function to return data table from SELECT query. 
    ' Param String sql, the SELECT statement. 
    ' Return DataTable of results.
    Public Function getDataTable(ByVal sql As String) As DataTable
        Dim dt As DataTable = New DataTable()
        Try
            Dim cnn As SQLiteConnection = New SQLiteConnection(dbConnection)
            cnn.Open()
            Dim command As SQLiteCommand = New SQLiteCommand(cnn)
            command.CommandText = sql
            Dim reader As SQLiteDataReader = command.ExecuteReader()
            dt.Load(reader)
            reader.Close()
            cnn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return dt
    End Function

    ' Function to insert data to the SQL database. 
    ' Param String table, the table to insert data into. 
    ' Param Dictionary(String, String), the columns and data to insert to.
    ' Return Boolean returnCode on success / fail. 
    Public Function insert(ByVal table As String, ByVal data As Dictionary(Of String, String)) As Boolean
        Dim columns As String = ""
        Dim values As String = ""
        Dim returnCode As Boolean = False

        For Each val As KeyValuePair(Of String, String) In data
            columns = columns & String.Format(val.Key.ToString() & ", ")
            values = values & String.Format("'" & val.Value & "', ")
        Next

        columns = columns.Substring(0, columns.Length - 2)
        values = values.Substring(0, values.Length - 2)

        Try
            returnCode = Me.executeQuery(String.Format("INSERT INTO " & table & "(" & columns & ") VALUES(" & values & ")"))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return returnCode
    End Function

    ' Function to execute a SQL statement. 
    ' Return Integer rowsUpdated - the number of rows affected by the query. 
    Public Function executeQuery(ByVal SQL As String)
        Dim cnn As New SQLiteConnection(dbConnection)
        cnn.Open()
        Dim command As New SQLiteCommand(cnn)
        command.CommandText = SQL
        Dim rowsUpdated As Integer = command.ExecuteNonQuery()
        cnn.Close()
        Return rowsUpdated
    End Function

#Region "IDisposable Support"

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' Dispose managed state (managed objects).
            End If

        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class
