Option Explicit On

Imports System.Diagnostics

Public Class CSAPScripts
    ' Class to contain all the SAP Scripts. The Ribbon currently have support to hold 30 different scripts. 
    ' The scripts here are called dynamically from the SAPMainScript function with the CallByName function. 
    ' Created by MOAN Enterprise 2015-06-25. Updated 2015-06-28. 

    ' Debug function example. Alter the contents here to a script that you want to debug. The actual scripts need to have the session as a variable. That is not needed here.
    ' Return String SAP Statusbar message (empty in this case). 
    Public Function EnterNetworkScript(ByVal args(,) As Object) As String

        ' Dimensions. 
        Dim var1, var2 As Object

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

    ' Function to update the Invoice date in an invoice. 
    ' Return String the text in the SAP Statusbar. 
    Public Function UpdateNetworkHeaderDateScript(ByVal args(,) As Object) As String

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

        Return session.findById("wnd[0]/sbar").Text

    End Function

End Class
