Option Explicit On

Public Class CTimer

    ' Class to encapsulate methods to handle the timer, responsible for keeping track of the progress left.

    Private timeAtLastProcessedObject As Double
    Private startTime As Double
    Private stopTime As Double
    Private averageTimeToNow As Double
    Private elapsedTimePerObject As Double

    ' Constructor function. 
    ' Return void. 
    Public Sub New()

    End Sub

    Public Sub start()
        Me.startTime = Now.ToOADate()
    End Sub

    ' Function to recalculate the time left of the macro.
    ' Param Integer objectsCountValue, the 
    ' Return void.
    Public Sub reCalculate(ByVal currentObject As Integer, ByVal objectsLeft As Integer)
        timeAtLastProcessedObject = Now.ToOADate()
        If Not objectsLeft = 0 Then
            Me.averageTimeToNow = (Me.timeAtLastProcessedObject - Me.startTime) / currentObject
            Me.elapsedTimePerObject = averageTimeToNow / currentObject
        End If
    End Sub

    ' Function to return the time remaining, based on the activecell.
    ' Return integer objects left to process.
    Public Function getTimeLeftMinutes(ByVal objectsLeft As Integer)
        Return Math.Round(((elapsedTimePerObject * 1400) * objectsLeft), 2)
    End Function

    ' Function to get the remaining time in seconds.
    ' Return double, remaining time of processing.
    Public Function getTimeLeftSeconds(ByVal objectsLeft As Integer)
        Return Math.Round(((averageTimeToNow * 86400) * objectsLeft) Mod 60, 0)
    End Function

    ' Function to stop the timer. Should be called when the loop is done.
    ' Return void.
    Public Sub stopTimer()
        stopTime = Now.ToOADate()
    End Sub

    ' Function to return the total elapsed time in seconds.
    ' Return double total time of excecution.
    Public Function getTotalTimeElapsedTimeInSeconds() As Integer
        Return (stopTime - startTime) * 86400
    End Function

End Class
