Option Explicit On

''' <summary>
''' A module for establishing the connection with SAP. The property "session" is used from SAPMain. 
''' </summary>
''' <remarks></remarks>
Module SAPConnection

    ' Static methods to to establish connection to SAP.
    Private app
    Private sapGuiAuto
    Private con
    Private WScript
    Private _session As Object
    Private xlApp As Excel.Application
    Private createdSession As Boolean
    Const MAX_SESSIONS As Integer = 6

    ''' <summary>
    ''' Encapsulate the _session variable, that is really a GuiSession object. 
    ''' </summary>
    ''' <value>The member _session.</value>
    ''' <returns>Private member _session.</returns>
    ''' <remarks></remarks>
    Public Property session() As Object
        Get
            Return _session
        End Get
        Private Set(ByVal value As Object)
            _session = value
        End Set
    End Property

    ''' <summary>
    ''' This function takes over a existing session from SAP. Basically the same as generated from SAP apart from some variable renaming. 
    ''' </summary>
    ''' <param name="transaction">The transaction to be entered before starting the object update loop.</param>
    ''' <remarks></remarks>
    Public Sub initWithTransaction(ByVal transaction As String)
        createdSession = False

        ' Reset any open connections before starting.
        resetSession()

        If IsNothing(app) Then
            sapGuiAuto = GetObject("SAPGUI")
            app = sapGuiAuto.GetScriptingEngine
        End If
        If IsNothing(con) Then
            con = app.Children(0)
        End If
        ' If Not IsObject(session) Then
        If IsNothing(_session) Then
            _session = con.Children(0)
        End If
        If Not IsNothing(WScript) Then
            WScript.ConnectObject(_session, "on")
            WScript.ConnectObject(xlApp, "on")
        End If

        ' Iconify the session to increase local performance. Inactivated when debugging!
        ' session.FindById("wnd[0]").Maximize()

        _session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n" & transaction
        _session.FindById("wnd[0]").sendVKey(0)
    End Sub

    ''' <summary>
    ''' This function creates a new GuiSession and takes over it, allowing the user to keep on working in SAP. Basically the same as generated from SAP apart from some variable renaming. 
    ''' </summary>
    ''' <param name="transaction">The transaction to be entered before starting the object update loop.</param>
    ''' <remarks></remarks>
    Public Sub initWithNewSessionAndTransaction(ByVal transaction As String)
        createdSession = False

        ' Reset any open connections before starting.
        resetSession()

        If IsNothing(app) Then
            sapGuiAuto = GetObject("SAPGUI")
            app = sapGuiAuto.GetScriptingEngine
        End If
        If IsNothing(con) Then
            con = app.Children(0)
        End If
        ' If Not IsObject(session) Then
        If IsNothing(_session) Then
            _session = con.Children(0)
        End If
        If Not IsNothing(WScript) Then
            WScript.ConnectObject(_session, "on")
            WScript.ConnectObject(xlApp, "on")
        End If

        Dim sessionCount As Integer
        sessionCount = con.Children.Count - 1
        Dim sessionNumber(0 To 6) As Integer

        Dim i As Integer
        For i = 1 To MAX_SESSIONS
            sessionNumber(i) = 0
        Next i

        For i = 0 To sessionCount
            _session = con.Children(Int(i))
            sessionNumber(_session.Info.sessionNumber) = _session.Info.sessionNumber
        Next i

        If (sessionCount < (MAX_SESSIONS - 1)) Then
            _session.CreateSession()
            Do
                ' no xl waits here preferably. 
                ' xlApp.Wait(Now + TimeValue("00:00:01"))
                If (con.Children.Count - sessionCount >= 2) Then Exit Do
            Loop
            On Error Resume Next
            Dim errNumb As Integer
            errNumb = 1
            For i = 0 To sessionCount + 1
                Err.Clear()
                _session = con.Children(Int(i))
                If (Err.Number > 0 Or Err.Number < 0) Then Exit For
                If (sessionNumber(_session.Info.sessionNumber) = 0) Then
                    errNumb = 0
                    Exit For
                End If
            Next i
            On Error GoTo 0
            createdSession = True
        Else
            MsgBox("You cannot open a new session. Most probably you have too many open. Close one and try again!", vbCritical, "Error")
        End If

        ' session.FindById("wnd[0]").Maximize()

        ' Iconify the session to increase local performance. Inactivated when debugging!
        ' session.FindById("wnd[0]").Iconify

        _session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n" & transaction
        _session.FindById("wnd[0]").sendVKey(0)

    End Sub

    ''' <summary>
    ''' If a session is created, close that window and null the objects. 
    ''' </summary>
    ''' <remarks>Does not really close the connection with SAP. This module needs to be nulled too?</remarks>
    Public Sub closeConnection()
        If createdSession = True Then
            _session.FindById("wnd[0]").Close()
        End If
        ' Barber pole is not disabled. 
        _session = Nothing
        con = Nothing
        app = Nothing
        sapGuiAuto = Nothing
    End Sub

    ''' <summary>
    ''' Reset the connection to SAP. 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub resetSession()
        _session = Nothing
        sapGuiAuto = Nothing
    End Sub

    ''' <summary>
    ''' Reset the transaction if the script does not work. Close a window and enter the transaction again. 
    ''' </summary>
    ''' <param name="transaction">The transaction to be re-entered.</param>
    ''' <remarks></remarks>
    Public Sub resetTransaction(ByVal transaction As String)
        If Not _session.findById("wnd[1]", False) Is Nothing Then
            _session.findById("wnd[1]").Close()
        End If
        _session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n" & transaction
        _session.FindById("wnd[0]").sendVKey(0)
    End Sub

End Module
