Option Explicit On

Public Class CTimer

    ' Class to encapsulate methods to handle the timer, responsible for keeping track of the progress left.

    Private timeAtLastProcessedObject As Double
    Private startTime As Double
    Private stopTime As Double
    Private averageTimeToNow As Double
    Private elapsedTimePerObject As Double
    Private objectCount As Integer

    Private xlApp As Excel.Application

    ' Constructor function. 
    ' Set the startTime. 
    ' Return void. 
    Public Sub New()
        xlApp = Globals.ThisAddIn.Application
        startTime = Now.ToOADate()
    End Sub

    ' Function to recalculate the time left of the macro.
    ' Return void.
    Public Sub reCalculate(ByVal objectsCountValue)
        timeAtLastProcessedObject = Now.ToOADate()
        If Not objectsCountValue = 0 Then
            averageTimeToNow = (timeAtLastProcessedObject - startTime) / objectsCountValue
        End If
    End Sub

    ' Function to return the time remaining, based on the activecell.
    ' Return integer objects left to process.
    Public Function getTimeLeftMinutes()
        getTimeLeftMinutes = Math.Round(((elapsedTimePerObject * 1400) * getObjectsLeft()), 2)
    End Function

    ' Function to return the number of objects left to process.
    ' Return integer objects left to process.
    Public Function getObjectsLeft() As Integer
        If (xlApp.ActiveCell.Offset(1, 0).Value2 = "") Then
            getObjectsLeft = 0
        Else
            getObjectsLeft = xlApp.Range(xlApp.ActiveCell, xlApp.ActiveCell.End(Excel.XlDirection.xlDown)).Count - 1
        End If
    End Function

    ' Function to get the remaining time in seconds.
    ' Return double, remaining time of processing.
    Public Function getTimeLeftSeconds()
        getTimeLeftSeconds = Math.Round(((averageTimeToNow * 86400) * getObjectsLeft()) Mod 60, 0)
    End Function

    ' Function to stop the timer. Should be called when the loop is done.
    ' Return void.
    Public Sub stopTimer(ByVal objectCountValue)
        objectCount = objectCountValue
        stopTime = Now.ToOADate()
    End Sub

    ' Function to return the total elapsed time in seconds.
    ' Return double total time of excecution.
    Public Function getTotalTimeElapsedTimeInSeconds() As Integer
        Return (stopTime - startTime) * 86400
    End Function

End Class
