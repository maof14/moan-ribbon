Imports Microsoft.Office.Tools.Ribbon
Imports System.Data.SQLite
Imports System.Data
Imports System.Diagnostics

Public Class Ribbon1

    ' This class acts as the old "Ribbon callbacks". Should be as little logic here as possible and store that in classes. 

    Private xlApp As Excel.Application
    Private scriptData As Dictionary(Of String, String)

    ' Event function for when the Ribbon loads. 
    ' Set the xlApp variable in this class. 
    ' Return void. 
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        xlApp = Globals.ThisAddIn.Application
        Debug.Print("Ribbon loaded successfully.")
        Me.grpVersion.Label = "Version " & My.Application.Info.Version.ToString()

    End Sub

    ' Event function for when the user clicks one of the 30 buttons in the Script menu. 
    ' Function loads information of the script from the SQLite database. 
    ' Return void. 
    Private Sub Button_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button1.Click, _
                                                                                                                                        Button2.Click, _
                                                                                                                                        Button3.Click, _
                                                                                                                                        Button4.Click, _
                                                                                                                                        Button5.Click, _
                                                                                                                                        Button6.Click, _
                                                                                                                                        Button7.Click, _
                                                                                                                                        Button8.Click, _
                                                                                                                                        Button9.Click, _
                                                                                                                                        Button10.Click, _
                                                                                                                                        Button11.Click, _
                                                                                                                                        Button12.Click, _
                                                                                                                                        Button13.Click, _
                                                                                                                                        Button14.Click, _
                                                                                                                                        Button15.Click, _
                                                                                                                                        Button17.Click, _
                                                                                                                                        Button18.Click, _
                                                                                                                                        Button19.Click, _
                                                                                                                                        Button20.Click, _
                                                                                                                                        Button21.Click, _
                                                                                                                                        Button22.Click, _
                                                                                                                                        Button23.Click, _
                                                                                                                                        Button24.Click, _
                                                                                                                                        Button25.Click, _
                                                                                                                                        Button26.Click, _
                                                                                                                                        Button27.Click, _
                                                                                                                                        Button28.Click, _
                                                                                                                                        Button29.Click, _
                                                                                                                                        Button30.Click
        Dim scriptName As String = sender.Label

        Dim db As CDatabase = New CDatabase()
        Dim res As DataTable = db.getDataTable("SELECT * FROM scripts WHERE name = '" & scriptName & "'")

        scriptData = New Dictionary(Of String, String)
        Dim i As Integer = 0

        For Each r In res.Rows
            scriptData.Add("id", r("id"))
            scriptData.Add("name", r("name"))
            scriptData.Add("description", r("description"))
            scriptData.Add("creator", r("creator"))
            scriptData.Add("created", r("created"))
            scriptData.Add("scriptid", r("scriptid"))
            scriptData.Add("transaction", r("transaction"))
            scriptData.Add("category", r("category"))
            scriptData.Add("headers", r("headers"))
            If IsDBNull(r("validation")) Then
                scriptData.Add("validation", "")
            Else
                scriptData.Add("validation", r("validation"))
            End If
        Next

        Dim tg As CTemplateGenerator = New CTemplateGenerator()
        tg.InitiateTemplate(scriptData)
        db = Nothing
        tg = Nothing

    End Sub

    ' Event function for when the user clicks the Run button. 
    ' Return void. 
    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnRun.Click

        ' Params from SQL database.
        ' Execute SAPMainScript, throwing in the transaction (ie CJ20N) and scriptid (ie EnterNetwork)
        SAPMainScript(scriptData)

    End Sub

    ' Event function for when the user clicks the View statistics button. 
    ' Get the statistics from the database as array, and insert them into a new workbook. 
    ' Return void. 
    Private Sub btnViewStatistics_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnViewStatistics.Click

        Dim s As New CStatistics()
        Dim arr(,) As String = s.getStatistics()
        Dim count As Integer = arr.GetLength(0)

        xlApp.Workbooks.Add()
        xlApp.Range("A1:G" & count).Value = arr
        s = Nothing

    End Sub

    ' Event function for when the user clicks the Convert to String button. 
    ' Utilize CTools class to get the data. 
    ' Return void. 
    Private Sub btnToString_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnToString.Click

        Dim t As CTools
        t = New CTools
        For Each c In xlApp.Selection
            c.value2 = t.convertToString(c.value2)
        Next
        t = Nothing

    End Sub

    ' Event function for when the user clicks the Settings button. 
    ' Load the SettingsDialog form. 
    ' Return void. 
    Private Sub btnSettings_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnSettings.Click

        Dim sd As New SettingsDialog()
        sd.ShowDialog()
        sd = Nothing

    End Sub

    ' Event function for when the user clicks the Get current date button. 
    ' Utilize the CTools class to get the data. 
    ' Return void. 
    Private Sub btnGetDate_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnGetDate.Click

        Dim dateFormat As String = My.Settings.DateFormat
        Dim t As New CTools
        For Each c In xlApp.Selection
            c.value2 = t.getDate(dateFormat)
        Next
        t = Nothing

    End Sub

End Class
