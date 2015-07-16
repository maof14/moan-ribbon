Option Explicit On

''' <summary>
''' A class to contain all of the scripts executing things in SAP. The methods in this class is called dynamically from the SAPMain module with function CallByName. 
''' </summary>
''' <remarks>This class should contain ALL scripts for all customers, while each Ribbon presents different scripts. Good idea?</remarks>
Partial Public Class CSAPScripts

    ''' <summary>
    ''' Enter a network, and get out. 
    ''' </summary>
    ''' <param name="args">The parameters from Excel, as array.</param>
    ''' <returns>The message from SAP Status bar on success.</returns>
    ''' <remarks>Demo function.</remarks>
    Public Function EnterNetworkScript(ByVal args(,) As Object) As String

        ' Dimensions. 
        Dim var1, var2 As String

        ' Declare the variables. 
        var1 = args(1, 1)
        var2 = args(1, 2)

        ' Useless stuff. 
        'session.FindById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "                23"

        ' Press open button.
        session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton("OPEN")

        ' Enter variable. 
        session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-AUFNR").Text = var1

        ' Useless stuff
        ' session.FindById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-AUFNR").SetFocus()
        ' session.FindById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-AUFNR").caretPosition = 8

        ' Press enter. 
        session.findById("wnd[1]").sendVKey(0)

        ' Update the network description
        ' session.findbyid("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = var2
        ' session.findbyid("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").caretPosition = 19
        ' session.findbyid("wnd[0]").sendVKey 0

        ' Press back button. 
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        ' Useless stuff. 
        ' session.FindById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "               23"

        ' Return string Status bar message. 
        Return session.findById("wnd[0]/sbar").Text

    End Function

    ''' <summary>
    ''' Update the Billing header date on Billing documents. 
    ''' </summary>
    ''' <param name="args">The parameters from Excel, as array.</param>
    ''' <returns>The message in SAP Status bar.</returns>
    ''' <remarks></remarks>
    Public Function UpdateBillingHeaderDateScript(ByVal args(,) As Object) As String

        ' Dimensions. 
        Dim var1, var2 As String

        ' Assign variables. 
        var1 = args(1, 1)
        var2 = args(1, 2)

        ' Input the invoice in the field. 
        session.findbyid("wnd[0]/usr/ctxtVBRK-VBELN").Text = var1

        ' Press enter.
        session.findbyid("wnd[0]").sendVKey(0)

        ' Press button to get to header. 
        session.findbyid("wnd[0]/usr/btnTC_HEAD").press()

        ' Update the date to a new one. 
        session.findbyid("wnd[0]/usr/tabsTABSTRIP_OVERVIEW/tabpKFDE/ssubSUBSCREEN_BODY:SAPMV60A:6105/ctxtVBRK-FKDAT").Text = var2

        ' Press Enter again. 
        session.findbyid("wnd[0]").sendVKey(0)

        ' Save the invoice. 
        session.findbyid("wnd[0]/tbar[0]/btn[11]").press()

        ' Return string - the statusbar message. 
        Return session.findById("wnd[0]/sbar").Text

    End Function

    ''' <summary>
    ''' Update the Sales order System status per Sales order item. 
    ''' </summary>
    ''' <param name="args">The parameters from Excel, as array.</param>
    ''' <returns>The message in SAP Status bar.</returns>
    ''' <remarks>Sufficient error handling?</remarks>
    Public Function UpdateSalesOrderSystemStatusScript(ByVal args(,) As Object) As String

        Dim var1, var2 As String

        var1 = args(1, 1)
        var2 = args(1, 2)

        ' Enter the Sales Order. 
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = var1
        session.findById("wnd[0]").sendVKey(0)

        ' Press enter on any window. 
        If Not session.findById("wnd[1]", False) Is Nothing Then
            Do
                session.findById("wnd[1]").sendVKey(0)
            Loop While Not session.findById("wnd[1]", False) Is Nothing
        End If

        ' If status bar displays this message, then: (should possibly be, if warning... then). To pick up every warning, and then return the warning. 
        If session.findById("wnd[0]/sbar").Text = "Fin ext cust and/or Selling BU missing, please check your entries!" Then
            session.findById("wnd[0]").sendVKey(0)
        End If

        ' Enter the first sales order item. 
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").SetFocus()
        session.findById("wnd[0]").sendVKey(2)

        ' Check which tab to enter.
        If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\11").Text = "Status" Then
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\11").Select()
        Else
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12").Select()
        End If

        ' Set up item loop. 
        Do
            If session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/ctxtVBAP-PSTYV").Text = "ZVCO" Or session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/ctxtVBAP-PSTYV").Text = "ZHSS" Then GoTo ContinueNextItem

            Dim currentStatuses As String
            currentStatuses = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4456/txtRV45A-STTXT").Text

            ' If some status is already or does not comply - go on. 
            If var2 = "Set TECO" Then
                If InStr(currentStatuses, "TECO") > 0 Then GoTo ContinueNextItem
            ElseIf var2 = "Remove TECO" Then ' Om man ska ta bort TECO eller CLSD ..
                If InStr(currentStatuses, "REL") > 0 Then
                    GoTo ContinueNextItem
                ElseIf InStr(currentStatuses, "CLSD") > 0 Then
                    ' ActiveCell.Offset(0, 3).value = "Found CLSD item"
                    GoTo ContinueNextItem
                End If
            ElseIf var2 = "Remove CLSD" Then
                If InStr(currentStatuses, "TECO") > 0 Or InStr(currentStatuses, "REL") > 0 Then GoTo ContinueNextItem
            ElseIf var2 = "Set CLSD" Then
                If InStr(currentStatuses, "CLSD") > 0 Then GoTo ContinueNextItem
            End If

            If InStr(currentStatuses, "NoMP") > 0 Then GoTo ContinueNextItem

            ' Check if the button to alter status is available. 
            If (session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4456/btnBT_STAE")) Is Nothing Then GoTo ContinueNextItem
            If var2 = "Set TECO" Then
                ' wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4456/btnBT_STAE, om man inte vill använda findByName... Lite farligt. 
                session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press()
                session.findById("wnd[1]/usr/btnFCODE_BTAB").press() ' original SÄTT teco
            ElseIf var2 = "Remove TECO" Then
                session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press()
                session.ActiveWindow.FindByName("FCODE_BUTA", "GuiButton").press() ' Ta BORT TECO
            ElseIf var2 = "Remove CLSD" Then
                session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press()
                session.ActiveWindow.FindByName("FCODE_BUAB", "GuiButton").press() ' Ta BORT CLSD
            ElseIf var2 = "Set CLSD" Then
                session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press()
                session.ActiveWindow.FindByName("FCODE_STAB", "GuiButton").press() ' Sätt CLSD
            ElseIf var2 = "Set FNBL" Then
                session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press()
                session.ActiveWindow.FindByName("FCODE_STEF", "GuiButton").press() ' Sätt CLSD
            End If

            ' Label: Continue to the next sales order item. 
ContinueNextItem:
            session.findById("wnd[0]/tbar[1]/btn[19]").press()

        Loop While session.findById("wnd[0]/sbar").Text <> "There are no more items to be displayed"

        ' Try to save. 
        session.findById("wnd[0]/tbar[0]/btn[11]").press()

        ' Lots of various error handling here previously. Should instead return the error message and quit processing the document. 
        If Not session.findById("wnd[1]", False) Is Nothing Then
            If session.findById("wnd[1]").Text = "Information" Then
                session.findById("wnd[1]").sendVKey(0)
            End If
        End If

        Return session.findById("wnd[0]/sbar").Text

    End Function

    Public Function checkTable(ByVal args(,) As Object) As String
        Dim var1, var2 As String

        var1 = args(1, 1)
        var2 = args(1, 2)

        Dim s As String = ""

        session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").Text = var1
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btnPUSHSHOW").press()
        s = session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03P-DDTEXT[7,0]").Text
        'session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03P-DDTEXT[7,0]").caretPosition = 0
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        Return s

    End Function

End Class
