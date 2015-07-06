<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SettingsDialog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.txtDbPath = New System.Windows.Forms.TextBox()
        Me.lblScriptContainerFilePath = New System.Windows.Forms.Label()
        Me.btnBrowseDb = New System.Windows.Forms.Button()
        Me.lblDateFormat = New System.Windows.Forms.Label()
        Me.txtDateFormat = New System.Windows.Forms.TextBox()
        Me.chbMailErrors = New System.Windows.Forms.CheckBox()
        Me.grpMails = New System.Windows.Forms.GroupBox()
        Me.txtRecipients = New System.Windows.Forms.TextBox()
        Me.lblRecipients = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.grpMails.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(277, 274)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'txtDbPath
        '
        Me.txtDbPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDbPath.Location = New System.Drawing.Point(118, 6)
        Me.txtDbPath.Name = "txtDbPath"
        Me.txtDbPath.Size = New System.Drawing.Size(229, 20)
        Me.txtDbPath.TabIndex = 1
        '
        'lblScriptContainerFilePath
        '
        Me.lblScriptContainerFilePath.AutoSize = True
        Me.lblScriptContainerFilePath.Location = New System.Drawing.Point(12, 9)
        Me.lblScriptContainerFilePath.Name = "lblScriptContainerFilePath"
        Me.lblScriptContainerFilePath.Size = New System.Drawing.Size(96, 13)
        Me.lblScriptContainerFilePath.TabIndex = 2
        Me.lblScriptContainerFilePath.Text = "Database file path:"
        '
        'btnBrowseDb
        '
        Me.btnBrowseDb.Location = New System.Drawing.Point(353, 4)
        Me.btnBrowseDb.Name = "btnBrowseDb"
        Me.btnBrowseDb.Size = New System.Drawing.Size(70, 23)
        Me.btnBrowseDb.TabIndex = 5
        Me.btnBrowseDb.Text = "Browse..."
        Me.btnBrowseDb.UseVisualStyleBackColor = True
        '
        'lblDateFormat
        '
        Me.lblDateFormat.AutoSize = True
        Me.lblDateFormat.Location = New System.Drawing.Point(12, 42)
        Me.lblDateFormat.Name = "lblDateFormat"
        Me.lblDateFormat.Size = New System.Drawing.Size(65, 13)
        Me.lblDateFormat.TabIndex = 7
        Me.lblDateFormat.Text = "Date format:"
        '
        'txtDateFormat
        '
        Me.txtDateFormat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDateFormat.Location = New System.Drawing.Point(118, 39)
        Me.txtDateFormat.Name = "txtDateFormat"
        Me.txtDateFormat.Size = New System.Drawing.Size(78, 20)
        Me.txtDateFormat.TabIndex = 8
        '
        'chbMailErrors
        '
        Me.chbMailErrors.AutoSize = True
        Me.chbMailErrors.Location = New System.Drawing.Point(6, 19)
        Me.chbMailErrors.Name = "chbMailErrors"
        Me.chbMailErrors.Size = New System.Drawing.Size(145, 17)
        Me.chbMailErrors.TabIndex = 9
        Me.chbMailErrors.Text = "Send e-mails about errors"
        Me.chbMailErrors.UseVisualStyleBackColor = True
        '
        'grpMails
        '
        Me.grpMails.Controls.Add(Me.lblRecipients)
        Me.grpMails.Controls.Add(Me.txtRecipients)
        Me.grpMails.Controls.Add(Me.chbMailErrors)
        Me.grpMails.Location = New System.Drawing.Point(15, 74)
        Me.grpMails.Name = "grpMails"
        Me.grpMails.Size = New System.Drawing.Size(200, 100)
        Me.grpMails.TabIndex = 10
        Me.grpMails.TabStop = False
        Me.grpMails.Text = "Error mails"
        '
        'txtRecipients
        '
        Me.txtRecipients.Location = New System.Drawing.Point(103, 43)
        Me.txtRecipients.Name = "txtRecipients"
        Me.txtRecipients.Size = New System.Drawing.Size(78, 20)
        Me.txtRecipients.TabIndex = 10
        '
        'lblRecipients
        '
        Me.lblRecipients.AutoSize = True
        Me.lblRecipients.Location = New System.Drawing.Point(6, 46)
        Me.lblRecipients.Name = "lblRecipients"
        Me.lblRecipients.Size = New System.Drawing.Size(60, 13)
        Me.lblRecipients.TabIndex = 11
        Me.lblRecipients.Text = "Recipients:"
        '
        'SettingsDialog
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(435, 315)
        Me.Controls.Add(Me.grpMails)
        Me.Controls.Add(Me.txtDateFormat)
        Me.Controls.Add(Me.lblDateFormat)
        Me.Controls.Add(Me.btnBrowseDb)
        Me.Controls.Add(Me.lblScriptContainerFilePath)
        Me.Controls.Add(Me.txtDbPath)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettingsDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Settings"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.grpMails.ResumeLayout(False)
        Me.grpMails.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents txtDbPath As System.Windows.Forms.TextBox
    Friend WithEvents lblScriptContainerFilePath As System.Windows.Forms.Label
    Friend WithEvents btnBrowseDb As System.Windows.Forms.Button
    Friend WithEvents lblDateFormat As System.Windows.Forms.Label
    Friend WithEvents txtDateFormat As System.Windows.Forms.TextBox
    Friend WithEvents chbMailErrors As System.Windows.Forms.CheckBox
    Friend WithEvents grpMails As System.Windows.Forms.GroupBox
    Friend WithEvents txtRecipients As System.Windows.Forms.TextBox
    Friend WithEvents lblRecipients As System.Windows.Forms.Label

End Class
