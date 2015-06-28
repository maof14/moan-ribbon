﻿Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.FileIO

Public Class SettingsDialog

    ' Event function for the when the user clicks the OK button. 
    ' Check if the files exist, and save the updates to the app settings. 
    ' Return void. 
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        If (File.Exists(txtDbPath.Text)) Then

            My.Settings.AppDbPath = txtDbPath.Text
            My.Settings.DateFormat = txtDateFormat.Text

            ' Perform the actual save of the settings.
            My.Settings.Save()
        Else
            MsgBox("Oops. Looks like the file on the specified path does not exist.", vbInformation, "Error in path")
        End If
        Me.Close()
    End Sub

    ' Event function for when the user clicks the Cancel button. 
    ' Do not do anything, just close the form. 
    ' Return void. 
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    ' Event function for when the user clicks the Browse button (database file). 
    ' Display a FileDialog. If the user chooses a file, set the path to the textfield corresponding to the button. 
    ' Return void. 
    Private Sub btnBrowseDb_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowseDb.Click
        Dim fd As New OpenFileDialog()
        fd.Title = "Choose the database file"
        fd.InitialDirectory = SpecialDirectories.MyDocuments
        fd.Filter = ".db3 files (*.db3*)|*.db3"
        fd.FilterIndex = 1
        fd.RestoreDirectory = True
        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.txtDbPath.Text = fd.FileName
        End If
    End Sub

    ' Constructor function. 
    ' Set the paths to the TextFields from the app settings. 
    ' Return void. 
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Me.txtDbPath.Text = My.Settings.AppDbPath
        Me.txtDateFormat.Text = My.Settings.DateFormat

    End Sub
End Class
