Option Explicit On

Imports System.Diagnostics

Module SAPMain
    ' Created by Mattias Olsson XB (qolsmat) 2014-04-08. Updated by 2015-06-02. Trying with VB.Net 2015-06-22
    ' Main function of the MOAN SAP Ribbon Add-In. Responsible for coordinating the classes used for SAP scripting and calling the Sub Scripts executing updates in SAP.
    ' Return void.

    ' This class should communicate directly with excel. Or should another class be used for that? 

    ' Members. 
    Private xlApp As Excel.Application

    Sub SAPMainScript(ByVal scriptData As Dictionary(Of String, String))

        ' Init the xlApp, "Application" in VBA. 
        xlApp = Globals.ThisAddIn.Application

        ' Dimension some variables to the SAPMainScript subprocedure.
        Dim stats As CStatistics
        Dim timer As CTimer
        Dim counter As Integer = 0
        Dim rng As Excel.Range
        Dim args(,) As Object

        ' Declare soma variables from scriptData. 
        Dim transaction As String = scriptData("transaction")
        Dim scriptid As String = scriptData("scriptid")

        ' Initialize dependency classes
        stats = New CStatistics()
        timer = New CTimer()

        ' Select the first cell to update in the template.
        xlApp.Cells(6, 2).Select()

        ' Try and initialize a session with transaction <transaction>
        Try
            initWithNewSessionAndTransaction(transaction)
        Catch ex As Exception
            MsgBox(ex.Message) ' Mattias-PC; Cannot initiate ActiveX Component. 
        End Try

        ' Init the script container class. To be able to call methods dynamically. 
        Dim ss As CSAPScripts = New CSAPScripts()

        ' Setup SAP object loop.
        ' Todo: Make it enter the second one too. 
        Do Until (IsNothing(xlApp.ActiveCell.Value))
            ' Take the arguments of the row.

            rng = xlApp.Range(xlApp.ActiveCell, xlApp.ActiveCell.End(Excel.XlDirection.xlToRight))
            args = rng.Value

            ' Execute the script.
            Try
                xlApp.Cells(xlApp.ActiveCell.Row, xlApp.ActiveCell.End(Excel.XlDirection.xlToRight).Column + 1).Value2 = CallByName(ss, scriptid, CallType.Method, args)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            ' Move down one cell and increment the counter by 1. 
            xlApp.ActiveCell.Offset(1, 0).Select()
            counter = counter + 1

            ' Recalculate time left and output in Excel statusbar.
            timer.reCalculate(counter)
            xlApp.StatusBar = timer.getTimeLeftMinutes() & " minutes left and " & timer.getTimeLeftSeconds() & " seconds. Objects left: " & timer.getObjectsLeft() & "."
        Loop

        ' Updated objects done - stop the timer.
        timer.stopTimer(counter)

        ' Create dictionary with statistics details. 

        Dim statisticsData As New Dictionary(Of String, String)

        statisticsData("scriptid") = scriptData("scriptid")
        statisticsData("objectcount") = counter
        statisticsData("username") = xlApp.UserName
        statisticsData("errorcount") = 0
        statisticsData("finished") = Now().ToLocalTime()
        statisticsData("finishedin") = timer.getTotalTimeElapsedTimeInSeconds()

        ' Insert the statistics in the database. 
        stats.writeStatistics(statisticsData)

        ' Done with stats, timer and scriptclass, dispose the objects.
        stats = Nothing
        timer = Nothing
        ss = Nothing

        ' Close the connection to SAP.
        closeConnection()

        ' Give Excel control over to the statusbar.
        xlApp.StatusBar = False

    End Sub

End Module
