Option Explicit On

''' <summary>
''' A class that handles timer functionality when executing the scripts. 
''' </summary>
''' <remarks></remarks>
Public Class CTimer

    Private timeAtLastProcessedObject As Double
    Private startTime As Double
    Private stopTime As Double
    Private averageTimeToNow As Double
    Private elapsedTimePerObject As Double

    ''' <summary>
    ''' Default constructor. Does nothing at the moment. 
    ''' </summary>
    ''' <remarks>Not doing anything.</remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' Starts the timer, before starting to execute the scripts. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub start()
        Me.startTime = Now.ToOADate()
    End Sub

    ''' <summary>
    ''' Recalculates the timer after eached updated object. 
    ''' </summary>
    ''' <param name="currentObject">The number of objects updated until now.</param>
    ''' <param name="objectsLeft">The total amount of objects to update.</param>
    ''' <remarks>Should be handled by event instead?</remarks>
    Public Sub reCalculate(ByVal currentObject As Integer, ByVal objectsLeft As Integer)
        timeAtLastProcessedObject = Now.ToOADate()
        If Not objectsLeft = 0 Then
            Me.averageTimeToNow = (Me.timeAtLastProcessedObject - Me.startTime) / currentObject
            Me.elapsedTimePerObject = averageTimeToNow / currentObject
        End If
    End Sub

    ''' <summary>
    ''' Return the calculated remaining time, based on the average time it took per object until now. 
    ''' </summary>
    ''' <param name="objectsLeft">The amount of objects left to update.</param>
    ''' <returns>The number of minutes left until the program is done.</returns>
    ''' <remarks>Added a return type here, without testing.</remarks>
    Public Function getTimeLeftMinutes(ByVal objectsLeft As Integer) As Double
        Return Math.Round(((elapsedTimePerObject * 1400) * objectsLeft), 2)
    End Function

    ''' <summary>
    ''' Return the calculated remaining time in seconds. To be used together with the minutes, becase this method excludes them.
    ''' </summary>
    ''' <param name="objectsLeft">The amount of objects left to update.</param>
    ''' <returns>The number of seconds left, minutes excluded.</returns>
    ''' <remarks>Added a return type here, without testing.</remarks>
    Public Function getTimeLeftSeconds(ByVal objectsLeft As Integer) As Integer
        Return Math.Round(((averageTimeToNow * 86400) * objectsLeft) Mod 60, 0)
    End Function

    ''' <summary>
    ''' Stops the timer. To be called after the main loop. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub stopTimer()
        stopTime = Now.ToOADate()
    End Sub

    ''' <summary>
    ''' Return the total elapsed time of executing the scripts. To be used when reporting statistics. 
    ''' </summary>
    ''' <returns>The total amount of time elapsed.</returns>
    ''' <remarks></remarks>
    Public Function getTotalTimeElapsedTimeInSeconds() As Integer
        Return (stopTime - startTime) * 86400
    End Function

End Class
