Public Class ProcessForm

    Private timer As CTimer

    ' Event function for when the user presses the Cancel processing button. 
    ' Set the DoWork property to false, causing the MainScript to quit after next saved object. 
    ' Return void. 
    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        SAPMain.DoWork = False
        Me.Close()
    End Sub

    ' Constructor function for the form. 
    ' Make the form appear on top of the other windows. 
    ' Return void. 
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        timer = New CTimer()
        timer.start()
    End Sub

    ''' <summary>
    ''' Function to update the progress bar. 
    ''' Param Integer currentItem, the position where the script is currently at. 
    ''' Param Integer totalItems, the total number of items to process. 
    ''' Return void. 
    ''' </summary>    
    Public Sub updateProgress(ByVal currentItem As Integer, ByVal totalItems As Integer)
        Me.timer.reCalculate(currentItem, totalItems)
        ' Todo: Fix "invisible" exception. 
        ' May need to implement BackgroundWorker. 
        ' http://stackoverflow.com/questions/18762673/why-the-cross-threading-exception-raises-only-when-debugging
        ' http://stackoverflow.com/questions/19596091/using-thread-to-open-form
        Me.lblTimeLeft.Text = "Estimated completion in " & Math.Floor(timer.getTimeLeftMinutes(totalItems - currentItem)) & " minutes and " & Math.Floor(timer.getTimeLeftSeconds(totalItems - currentItem)) & " seconds."
        Me.prgProgress.Value = Math.Floor(currentItem / totalItems)
        Me.prgProgress.Refresh()
    End Sub

    ' Function to get results to somewhere. Ie statusbar. 
    ' Todo: Implement this function. 
    Public Function getResults()
        Me.timer.stopTimer()
        Return Me.timer.getTotalTimeElapsedTimeInSeconds()
    End Function

    ' Function to finalize the class. 
    ' Disposes the timer. 
    ' Return void. 
    Protected Overrides Sub Finalize()
        Me.timer.Dispose()
        MyBase.Finalize()
    End Sub
End Class