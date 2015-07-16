Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SQLite
Imports System.IO

''' <summary>
''' A database wrapper class for use with SQLite.
''' </summary>
''' <remarks></remarks>
Public Class CSQLiteDatabase : Implements IDatabase

    ' Database wrapper for SQLite3 (ADO)
    ' @see more CRUD http://www.dreamincode.net/forums/topic/157830-using-sqlite-with-c%23/
    ' ADO.DB connectors for System.Data.SQLite: https://system.data.sqlite.org/index.html/doc/trunk/www/downloads.wiki

    Private dbConnection As String

    ''' <summary>
    ''' Basic constructor, setting up a standard connection string to the database with the db in the settings. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Dim dbPath As String = My.Settings.AppDbPath
        Me.dbConnection = "Data Source=" & dbPath
    End Sub

    ''' <summary>
    ''' Overloading constructor, setting up the connection string to a custom. 
    ''' </summary>
    ''' <param name="dbFullPath">A path to the database.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbFullPath As String)
        Me.dbConnection = "Data Source=" & dbFullPath
    End Sub

    ''' <summary>
    ''' Select data from the database, returning a datatable with the results. 
    ''' </summary>
    ''' <param name="sql">The SELECT query.</param>
    ''' <returns>DataTable with the results.</returns>
    ''' <remarks></remarks>
    Public Function getDataTable(ByVal sql As String) As DataTable Implements IDatabase.getDataTable
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

    ''' <summary>
    ''' Insert data to the database. 
    ''' </summary>
    ''' <param name="table">The table to be affected by the query.</param>
    ''' <param name="data">The data to be inserted, column and value.</param>
    ''' <returns>Boolean indicating success.</returns>
    ''' <remarks></remarks>
    Public Function insert(ByVal table As String, ByVal data As Dictionary(Of String, String)) As Boolean Implements IDatabase.insert
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

    ''' <summary>
    ''' Execute a query, that returns no data. 
    ''' </summary>
    ''' <param name="sql">The SQL statement.</param>
    ''' <returns>Integer on success</returns>
    ''' <remarks></remarks>
    Public Function executeQuery(ByVal SQL As String) As Integer Implements IDatabase.executeQuery
        Dim cnn As New SQLiteConnection(dbConnection)
        cnn.Open()
        Dim command As New SQLiteCommand(cnn)
        command.CommandText = SQL
        Dim rowsUpdated As Integer = command.ExecuteNonQuery()
        cnn.Close()
        Return rowsUpdated
    End Function

    ' To be implemented... 

    ''' <summary>
    ''' Delete data from the database.
    ''' </summary>
    ''' <param name="table">The table to be affected by the query.</param>
    ''' <param name="data">The data to be deleted, column and value.</param>
    ''' <returns>Boolena indicating success.</returns>
    ''' <remarks></remarks>
    Public Function delete(ByVal table As String, ByVal data As System.Collections.Generic.Dictionary(Of String, String)) As Boolean Implements IDatabase.delete
        Return False
    End Function

    ''' <summary>
    ''' Update data in the database.
    ''' </summary>
    ''' <param name="table">The table to be affected by the query.</param>
    ''' <param name="data">The data to be updated, column and value.</param>
    ''' <returns>Boolean indicating success.</returns>
    ''' <remarks></remarks>
    Public Function update(ByVal table As String, ByVal data As Dictionary(Of String, String)) As Boolean Implements IDatabase.update
        Return False
    End Function

End Class
