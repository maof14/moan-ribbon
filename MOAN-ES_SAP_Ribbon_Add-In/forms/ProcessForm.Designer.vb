<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProcessForm
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
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lblCancelInfo = New System.Windows.Forms.Label()
        Me.prgProgress = New System.Windows.Forms.ProgressBar()
        Me.lblTimeLeft = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblInfo
        '
        Me.lblInfo.AutoSize = True
        Me.lblInfo.Font = New System.Drawing.Font("Trebuchet MS", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfo.Location = New System.Drawing.Point(28, 9)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(500, 46)
        Me.lblInfo.TabIndex = 0
        Me.lblInfo.Text = "The script is now processing."
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(195, 227)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(148, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel processing"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblCancelInfo
        '
        Me.lblCancelInfo.AutoSize = True
        Me.lblCancelInfo.Location = New System.Drawing.Point(87, 200)
        Me.lblCancelInfo.Name = "lblCancelInfo"
        Me.lblCancelInfo.Size = New System.Drawing.Size(384, 13)
        Me.lblCancelInfo.TabIndex = 2
        Me.lblCancelInfo.Text = "If you cancel the processing, the script will terminate after the next updated it" & _
    "em."
        '
        'prgProgress
        '
        Me.prgProgress.Location = New System.Drawing.Point(36, 89)
        Me.prgProgress.Name = "prgProgress"
        Me.prgProgress.Size = New System.Drawing.Size(492, 23)
        Me.prgProgress.TabIndex = 3
        '
        'lblTimeLeft
        '
        Me.lblTimeLeft.AutoSize = True
        Me.lblTimeLeft.Location = New System.Drawing.Point(170, 151)
        Me.lblTimeLeft.Name = "lblTimeLeft"
        Me.lblTimeLeft.Size = New System.Drawing.Size(205, 13)
        Me.lblTimeLeft.TabIndex = 4
        Me.lblTimeLeft.Text = "Waiting for the first object to be updated..."
        '
        'ProcessForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(555, 262)
        Me.Controls.Add(Me.lblTimeLeft)
        Me.Controls.Add(Me.prgProgress)
        Me.Controls.Add(Me.lblCancelInfo)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblInfo)
        Me.Name = "ProcessForm"
        Me.Text = "Processing"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblCancelInfo As System.Windows.Forms.Label
    Friend WithEvents prgProgress As System.Windows.Forms.ProgressBar
    Friend WithEvents lblTimeLeft As System.Windows.Forms.Label
End Class
