Option Explicit On

' Class to encapsulate the handling of the writing and reading of the statistics database.
' Created by MOAN Enterprise 2015-06-27. 

Public Class CStatistics

    Private db As CDatabase

    Public Sub New()
        Me.db = New CDatabase()
    End Sub

    ' Function to write the statistics to the database.
    ' This method should be run after all SAP execution is complete in SAPMacro.
    ' Return Boolean success from CDatabase.
    Public Function writeStatistics(ByVal scriptData As Dictionary(Of String, String)) As Boolean
        Return Me.db.insert("statistics", scriptData)
    End Function

    ' Function to get the statistics data for display. 
    ' Todo: Add functionality to display statistics. 
    ' Return statistics data as some data type. 
    Public Function getStatistics() As Integer
        Return 0
    End Function

End Class
