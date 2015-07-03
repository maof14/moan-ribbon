﻿Option Explicit On

Imports System.Data
Imports System.ComponentModel

' Class to encapsulate the handling of the writing and reading of the statistics database.
' Created by MOAN Enterprise 2015-06-27. 

Public Class CStatistics : Implements IDisposable

    Private db As CDatabase
    Private disposedValue As Boolean ' To detect redundant calls

    ''' <summary>
    ''' Constructor for class CStatistics. Initializes the CDatabase member. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Me.db = New CDatabase()
    End Sub

    ' Function to write the statistics to the database.
    ' This method should be run after all SAP execution is complete in SAPMacro.
    ' Return Boolean success from CDatabase.
    Public Function writeStatistics(ByVal scriptData As Dictionary(Of String, String)) As Boolean
        Return Me.db.insert("statistics", scriptData)
    End Function

    ' Function to get the statistics data for display from the database. 
    ' Should there be a function to have this to display Pivot table too? How to reference the Pivot in creation, and choose layout?
    ' Return String array of the database statistics results. 
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

#Region "IDisposable Support"

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
