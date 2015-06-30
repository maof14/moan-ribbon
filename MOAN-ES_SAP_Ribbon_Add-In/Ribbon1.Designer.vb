Imports Microsoft.Office.Tools.Ribbon
Imports System.Data.SQLite
Imports System.Data
Imports System.Diagnostics

Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

        ' Init the available scripts in the Ribbon... Perhaps this should be somewhere else, with Menu1.Dynamic = True, if possible. 
        ' Fetch the available scripts from the db located at dbPath. Should be located on network drive. 

        Dim db As CDatabase = New CDatabase()
        Dim res As DataTable = db.getDataTable("SELECT * FROM scripts ORDER BY category ASC")

        Dim i As Integer = 0
        Dim category As String = ""
        Dim btn As RibbonButton
        Dim sep As RibbonSeparator

        ' Update the buttons in Menu1 (name?). 30 buttons are available as default. Add more if needed. 
        For Each r In res.Rows
            If r("category") <> category Then
                category = r("category")
                sep = Globals.Factory.GetRibbonFactory().CreateRibbonSeparator()
                sep.Title = category
                Menu1.Items.Insert(i, sep)
                i = i + 1
            End If
            btn = Menu1.Items(i)
            btn.Label = r("name")
            btn.ScreenTip = r("description")
            btn.Visible = True
            i = i + 1
        Next
    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tabMoan = Me.Factory.CreateRibbonTab
        Me.grpSAPAutomation = Me.Factory.CreateRibbonGroup
        Me.grpTools = Me.Factory.CreateRibbonGroup
        Me.grpStatistics = Me.Factory.CreateRibbonGroup
        Me.grpVersion = Me.Factory.CreateRibbonGroup
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Button22 = Me.Factory.CreateRibbonButton
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Button26 = Me.Factory.CreateRibbonButton
        Me.Button27 = Me.Factory.CreateRibbonButton
        Me.Button28 = Me.Factory.CreateRibbonButton
        Me.Button29 = Me.Factory.CreateRibbonButton
        Me.Button30 = Me.Factory.CreateRibbonButton
        Me.btnRun = Me.Factory.CreateRibbonButton
        Me.btnToString = Me.Factory.CreateRibbonButton
        Me.btnGetDate = Me.Factory.CreateRibbonButton
        Me.btnViewStatistics = Me.Factory.CreateRibbonButton
        Me.btnSettings = Me.Factory.CreateRibbonButton
        Me.btnCustom = Me.Factory.CreateRibbonButton
        Me.tabMoan.SuspendLayout()
        Me.grpSAPAutomation.SuspendLayout()
        Me.grpTools.SuspendLayout()
        Me.grpStatistics.SuspendLayout()
        Me.grpVersion.SuspendLayout()
        '
        'tabMoan
        '
        Me.tabMoan.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.tabMoan.Groups.Add(Me.grpSAPAutomation)
        Me.tabMoan.Groups.Add(Me.grpTools)
        Me.tabMoan.Groups.Add(Me.grpStatistics)
        Me.tabMoan.Groups.Add(Me.grpVersion)
        Me.tabMoan.Label = "MOAN ES"
        Me.tabMoan.Name = "tabMoan"
        '
        'grpSAPAutomation
        '
        Me.grpSAPAutomation.Items.Add(Me.Menu1)
        Me.grpSAPAutomation.Items.Add(Me.btnRun)
        Me.grpSAPAutomation.Label = "SAP automation"
        Me.grpSAPAutomation.Name = "grpSAPAutomation"
        '
        'grpTools
        '
        Me.grpTools.Items.Add(Me.btnToString)
        Me.grpTools.Items.Add(Me.btnGetDate)
        Me.grpTools.Label = "Tools"
        Me.grpTools.Name = "grpTools"
        '
        'grpStatistics
        '
        Me.grpStatistics.Items.Add(Me.btnViewStatistics)
        Me.grpStatistics.Label = "Statistics"
        Me.grpStatistics.Name = "grpStatistics"
        '
        'grpVersion
        '
        Me.grpVersion.Items.Add(Me.btnSettings)
        Me.grpVersion.Label = "Version 0.0"
        Me.grpVersion.Name = "grpVersion"
        '
        'Menu1
        '
        Me.Menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Items.Add(Me.Button1)
        Me.Menu1.Items.Add(Me.Button2)
        Me.Menu1.Items.Add(Me.Button3)
        Me.Menu1.Items.Add(Me.Button4)
        Me.Menu1.Items.Add(Me.Button5)
        Me.Menu1.Items.Add(Me.Button6)
        Me.Menu1.Items.Add(Me.Button7)
        Me.Menu1.Items.Add(Me.Button8)
        Me.Menu1.Items.Add(Me.Button9)
        Me.Menu1.Items.Add(Me.Button10)
        Me.Menu1.Items.Add(Me.Button11)
        Me.Menu1.Items.Add(Me.Button12)
        Me.Menu1.Items.Add(Me.Button13)
        Me.Menu1.Items.Add(Me.Button14)
        Me.Menu1.Items.Add(Me.Button15)
        Me.Menu1.Items.Add(Me.Button16)
        Me.Menu1.Items.Add(Me.Button17)
        Me.Menu1.Items.Add(Me.Button18)
        Me.Menu1.Items.Add(Me.Button19)
        Me.Menu1.Items.Add(Me.Button20)
        Me.Menu1.Items.Add(Me.Button21)
        Me.Menu1.Items.Add(Me.Button22)
        Me.Menu1.Items.Add(Me.Button23)
        Me.Menu1.Items.Add(Me.Button24)
        Me.Menu1.Items.Add(Me.Button25)
        Me.Menu1.Items.Add(Me.Button26)
        Me.Menu1.Items.Add(Me.Button27)
        Me.Menu1.Items.Add(Me.Button28)
        Me.Menu1.Items.Add(Me.Button29)
        Me.Menu1.Items.Add(Me.Button30)
        Me.Menu1.Label = "Create script template"
        Me.Menu1.Name = "Menu1"
        Me.Menu1.OfficeImageId = "PivotExportToExcel"
        Me.Menu1.ScreenTip = "Create a new script template."
        Me.Menu1.ShowImage = True
        '
        'Button1
        '
        Me.Button1.Label = "Button1"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        Me.Button1.Visible = False
        '
        'Button2
        '
        Me.Button2.Label = "Button2"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        Me.Button2.Visible = False
        '
        'Button3
        '
        Me.Button3.Label = "Button3"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        Me.Button3.Visible = False
        '
        'Button4
        '
        Me.Button4.Label = "Button4"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        Me.Button4.Visible = False
        '
        'Button5
        '
        Me.Button5.Label = "Button5"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        Me.Button5.Visible = False
        '
        'Button6
        '
        Me.Button6.Label = "Button6"
        Me.Button6.Name = "Button6"
        Me.Button6.ShowImage = True
        Me.Button6.Visible = False
        '
        'Button7
        '
        Me.Button7.Label = "Button7"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        Me.Button7.Visible = False
        '
        'Button8
        '
        Me.Button8.Label = "Button8"
        Me.Button8.Name = "Button8"
        Me.Button8.ShowImage = True
        Me.Button8.Visible = False
        '
        'Button9
        '
        Me.Button9.Label = "Button9"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        Me.Button9.Visible = False
        '
        'Button10
        '
        Me.Button10.Label = "Button10"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
        Me.Button10.Visible = False
        '
        'Button11
        '
        Me.Button11.Label = "Button11"
        Me.Button11.Name = "Button11"
        Me.Button11.ShowImage = True
        Me.Button11.Visible = False
        '
        'Button12
        '
        Me.Button12.Label = "Button12"
        Me.Button12.Name = "Button12"
        Me.Button12.ShowImage = True
        Me.Button12.Visible = False
        '
        'Button13
        '
        Me.Button13.Label = "Button13"
        Me.Button13.Name = "Button13"
        Me.Button13.ShowImage = True
        Me.Button13.Visible = False
        '
        'Button14
        '
        Me.Button14.Label = "Button14"
        Me.Button14.Name = "Button14"
        Me.Button14.ShowImage = True
        Me.Button14.Visible = False
        '
        'Button15
        '
        Me.Button15.Label = "Button15"
        Me.Button15.Name = "Button15"
        Me.Button15.ShowImage = True
        Me.Button15.Visible = False
        '
        'Button16
        '
        Me.Button16.Label = "Button16"
        Me.Button16.Name = "Button16"
        Me.Button16.ShowImage = True
        Me.Button16.Visible = False
        '
        'Button17
        '
        Me.Button17.Label = "Button17"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        Me.Button17.Visible = False
        '
        'Button18
        '
        Me.Button18.Label = "Button18"
        Me.Button18.Name = "Button18"
        Me.Button18.ShowImage = True
        Me.Button18.Visible = False
        '
        'Button19
        '
        Me.Button19.Label = "Button19"
        Me.Button19.Name = "Button19"
        Me.Button19.ShowImage = True
        Me.Button19.Visible = False
        '
        'Button20
        '
        Me.Button20.Label = "Button20"
        Me.Button20.Name = "Button20"
        Me.Button20.ShowImage = True
        Me.Button20.Visible = False
        '
        'Button21
        '
        Me.Button21.Label = "Button21"
        Me.Button21.Name = "Button21"
        Me.Button21.ShowImage = True
        Me.Button21.Visible = False
        '
        'Button22
        '
        Me.Button22.Label = "Button22"
        Me.Button22.Name = "Button22"
        Me.Button22.ShowImage = True
        Me.Button22.Visible = False
        '
        'Button23
        '
        Me.Button23.Label = "Button23"
        Me.Button23.Name = "Button23"
        Me.Button23.ShowImage = True
        Me.Button23.Visible = False
        '
        'Button24
        '
        Me.Button24.Label = "Button24"
        Me.Button24.Name = "Button24"
        Me.Button24.ShowImage = True
        Me.Button24.Visible = False
        '
        'Button25
        '
        Me.Button25.Label = "Button25"
        Me.Button25.Name = "Button25"
        Me.Button25.ShowImage = True
        Me.Button25.Visible = False
        '
        'Button26
        '
        Me.Button26.Label = "Button26"
        Me.Button26.Name = "Button26"
        Me.Button26.ShowImage = True
        Me.Button26.Visible = False
        '
        'Button27
        '
        Me.Button27.Label = "Button27"
        Me.Button27.Name = "Button27"
        Me.Button27.ShowImage = True
        Me.Button27.Visible = False
        '
        'Button28
        '
        Me.Button28.Label = "Button28"
        Me.Button28.Name = "Button28"
        Me.Button28.ShowImage = True
        Me.Button28.Visible = False
        '
        'Button29
        '
        Me.Button29.Label = "Button29"
        Me.Button29.Name = "Button29"
        Me.Button29.ShowImage = True
        Me.Button29.Visible = False
        '
        'Button30
        '
        Me.Button30.Label = "Button30"
        Me.Button30.Name = "Button30"
        Me.Button30.ShowImage = True
        Me.Button30.Visible = False
        '
        'btnRun
        '
        Me.btnRun.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnRun.Label = "Run script"
        Me.btnRun.Name = "btnRun"
        Me.btnRun.OfficeImageId = "MacroPlay"
        Me.btnRun.ScreenTip = "Run the current script template."
        Me.btnRun.ShowImage = True
        '
        'btnToString
        '
        Me.btnToString.Label = "Convert to string"
        Me.btnToString.Name = "btnToString"
        Me.btnToString.OfficeImageId = "InkToTextMode"
        Me.btnToString.ScreenTip = "Convert the selected values to string (prepend a apostrophe)."
        Me.btnToString.ShowImage = True
        '
        'btnGetDate
        '
        Me.btnGetDate.Label = "Get current date"
        Me.btnGetDate.Name = "btnGetDate"
        Me.btnGetDate.OfficeImageId = "GotoCalendar"
        Me.btnGetDate.ScreenTip = "Get todays date in the format specified in the settings."
        Me.btnGetDate.ShowImage = True
        '
        'btnViewStatistics
        '
        Me.btnViewStatistics.Label = "View statistics"
        Me.btnViewStatistics.Name = "btnViewStatistics"
        Me.btnViewStatistics.OfficeImageId = "AdpOutputOperationsAddToOutput"
        Me.btnViewStatistics.ScreenTip = "Lets you examine the statistics data in a new workbook."
        Me.btnViewStatistics.ShowImage = True
        '
        'btnSettings
        '
        Me.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSettings.Label = "Settings"
        Me.btnSettings.Name = "btnSettings"
        Me.btnSettings.OfficeImageId = "ToolboxGallery"
        Me.btnSettings.ScreenTip = "Open the settings dialog."
        Me.btnSettings.ShowImage = True
        '
        'btnCustom
        '
        Me.btnCustom.Label = "Custom scripts"
        Me.btnCustom.Name = "btnCustom"
        Me.btnCustom.OfficeImageId = "AdpOutputOperationsAddToOutput"
        Me.btnCustom.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.tabMoan)
        Me.tabMoan.ResumeLayout(False)
        Me.tabMoan.PerformLayout()
        Me.grpSAPAutomation.ResumeLayout(False)
        Me.grpSAPAutomation.PerformLayout()
        Me.grpTools.ResumeLayout(False)
        Me.grpTools.PerformLayout()
        Me.grpStatistics.ResumeLayout(False)
        Me.grpStatistics.PerformLayout()
        Me.grpVersion.ResumeLayout(False)
        Me.grpVersion.PerformLayout()

    End Sub

    Friend WithEvents tabMoan As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpSAPAutomation As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnRun As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpStatistics As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnViewStatistics As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpVersion As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnToString As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnGetDate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSettings As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button26 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button27 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button28 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button29 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button30 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCustom As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
