﻿Option Explicit On

Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Threading
Imports System.ComponentModel

''' <summary>
''' A module that contains the main loop of the program. 
''' </summary>
''' <remarks></remarks>
Module SAPMain
    ' Created by MOAN Enterprise. (qolsmat) 2014-04-08. Updated by 2015-06-02. Trying with VB.Net 2015-06-22

    ' Members. 
    Private xlApp As Excel.Application
    Private _doWork As Boolean = True
    Private pf As ProcessForm

    ' Properties. 
    ''' <summary>
    ''' Indicator for if the script shall continue. 
    ''' </summary>
    ''' <value>Boolean</value>
    ''' <returns>Boolean member _doWork</returns>
    ''' <remarks></remarks>
    Public Property DoWork() As Boolean
        Get
            Return _doWork
        End Get
        Set(ByVal value As Boolean)
            _doWork = value
        End Set
    End Property

    ' Methods. 
    
    ''' <summary>
    ''' Main program function. Responsible for coordinating all the classes in the program, communicating with Excel and running the scripts in SAP. 
    ''' </summary>
    ''' <param name="scriptData">A dictionary from the database, containing information about the script to be run.</param>
    ''' <remarks></remarks>
    Sub SAPMainScript(ByVal scriptData As Dictionary(Of String, String))

        ' Set DoWork to true each time the script start. Else if set to false, that persists in memory. 
        DoWork = True

        ' Init the xlApp, "Application" in VBA. 
        xlApp = Globals.ThisAddIn.Application

        ' Dimension some variables to the SAPMainScript subprocedure.
        Dim stats As CStatistics
        Dim em As CErrorMail
        Dim counter As Integer = 0
        Dim errorCounter As Integer = 0
        Dim rng As Excel.Range
        Dim args(,) As Object
        Dim objectsToUpdate As Integer = 0

        ' Declare soma variables from scriptData. 
        Dim transaction As String = scriptData("transaction")
        Dim scriptid As String = scriptData("scriptid")

        ' Initialize dependency classes
        stats = New CStatistics()
        em = New CErrorMail()
        em.recipient = My.Settings.ErrorRecipients

        ' Select the first cell to update in the template.
        xlApp.Cells(6, 2).Select()

        ' Count the total number of objects, for use with the progress handling.
        objectsToUpdate = xlApp.Range(xlApp.ActiveCell, xlApp.ActiveCell.End(Excel.XlDirection.xlDown)).Count

        ' Try and initialize a session with transaction <transaction>
        Try
            initWithNewSessionAndTransaction(transaction)
        Catch ex As Exception
            MsgBox(ex.Message) ' Mattias-PC; Cannot initiate ActiveX Component. 
        End Try

        ' Init the script container class. To be able to call methods dynamically. 
        Dim ss As CSAPScripts = New CSAPScripts()

        Dim t As Thread = New Thread(AddressOf ShowForm)
        t.SetApartmentState(ApartmentState.STA)
        t.Start()

        ' Setup SAP object loop.
        Do Until (IsNothing(xlApp.ActiveCell.Value))
            ' Take the arguments of the row.

            rng = xlApp.Range(xlApp.ActiveCell, xlApp.ActiveCell.End(Excel.XlDirection.xlToRight))
            args = rng.Value

            ' Execute the script.
            Try
                If DoWork = True Then
                    ' Must have the object to update, and one parameter to. (That is, at least two columns to fill in, and another for SAP Output.) 
                    xlApp.Cells(xlApp.ActiveCell.Row, xlApp.ActiveCell.End(Excel.XlDirection.xlToRight).Column + 1).Value2 = CallByName(ss, scriptid, CallType.Method, args)
                Else
                    Exit Do
                End If
            Catch ex As Exception
                errorCounter = errorCounter + 1
                em.addError(args(1, 1), scriptData("scriptid"), ex.Message, ex.Source)
                resetTransaction(scriptData("transaction"))
                Continue Do
            End Try

            ' Move down one cell and increment the counter by 1. 
            xlApp.ActiveCell.Offset(1, 0).Select()
            counter = counter + 1

            ' Update the ProgressForm progressbar and time left label with BeginInvoke. As the form is in another thread. 
            pf.BeginInvoke(New System.Action(Sub() pf.updateProgress(counter, objectsToUpdate)))
            pf.BeginInvoke(New System.Action(Sub() pf.worker.ReportProgress(CInt((counter / objectsToUpdate) * 100))))
        Loop

        ' Create dictionary with statistics details. 
        Dim statisticsData As New Dictionary(Of String, String)

        statisticsData("scriptid") = scriptData("scriptid")
        statisticsData("objectcount") = counter
        statisticsData("username") = xlApp.UserName
        statisticsData("errorcount") = errorCounter
        statisticsData("finished") = Now().ToLocalTime()
        statisticsData("finishedin") = pf.getResults()

        ' Insert the statistics in the database. 
        stats.writeStatistics(statisticsData)

        ' Check if there are any errors, and that the user want error mails. If not, don't do anything more with CErrorMails class. 
        If (errorCounter > 0 And My.Settings.ErrorMails = True) Then
            em.sendMail()
        End If

        ' Done with stats and scriptclass, throw away the objects (removing the pointers). 
        stats = Nothing
        ss = Nothing
        em = Nothing

        ' Close the window and connection to SAP.
        CloseConnection()

        ' Close the progress form, and send to nothing. 
        pf.BeginInvoke(New System.Action(Sub() pf.Dispose()))

        ' Give Excel control over to the statusbar.
        xlApp.StatusBar = False

    End Sub

    ''' <summary>
    ''' Function to handle the display of the process form that allows to cancel the script. 
    ''' Creates the progress form. To be handled in another thread. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ShowForm()
        pf = New ProcessForm
        Application.Run(pf)
    End Sub

End Module
