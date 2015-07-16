Option Explicit On

Imports System.Data
Imports System.ComponentModel

''' <summary>
''' Class to handle the statistics in the database. 
''' </summary>
''' <remarks></remarks>
Public Class CStatistics

    Private db As CSQLiteDatabase

    ''' <summary>
    ''' Default constructor. Initializes the database class. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Me.db = New CSQLiteDatabase()
    End Sub

    ''' <summary>
    ''' Write statistics to the database. 
    ''' </summary>
    ''' <param name="scriptData">The data to write to the database.</param>
    ''' <returns>Boolean indicating success.</returns>
    ''' <remarks></remarks>
    Public Function writeStatistics(ByVal scriptData As Dictionary(Of String, String)) As Boolean
        Return Me.db.insert("statistics", scriptData)
    End Function

    ''' <summary>
    ''' Get the statistics data from the database.
    ''' </summary>
    ''' <returns>The statistics as an array.</returns>
    ''' <remarks>Should be able to present as Pivot, too?</remarks>
    Public Function getStatistics() As String(,)
        Dim res As DataTable = Me.db.getDataTable("SELECT * FROM statistics")
        Dim count As Integer
        count = res.Rows.Count

        Dim arr(count + 1, 6) As String

        arr(0, 0) = "Id"
        arr(0, 1) = "Script id"
        arr(0, 2) = "Object count"
        arr(0, 3) = "Username"
        arr(0, 4) = "Error count"
        arr(0, 5) = "Finished"
        arr(0, 6) = "Finished in (seconds)"

        Dim i As Integer = 1

        For Each r In res.Rows
            arr(i, 0) = r("id")
            arr(i, 1) = r("scriptid")
            arr(i, 2) = r("objectcount")
            arr(i, 3) = r("username")
            arr(i, 4) = r("errorcount")
            arr(i, 5) = r("finished")
            arr(i, 6) = r("finishedin")
            i = i + 1
        Next

        Return arr
    End Function

End Class
