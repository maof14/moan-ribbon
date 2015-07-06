Option Explicit On

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

    ' session Property, to ensure encapsulation of _session. 
    ' Return GuiSession _session as Object. 
    ' Private set GuiSession _session as Object. 
    Public Property session() As Object
        Get
            Return _session
        End Get
        Private Set(ByVal value As Object)
            _session = value
        End Set
    End Property

    ' Function to take over a transaction from SAP. Basically the same as generated from SAP, but simplified and renamed due to duplicate naming with Excel pre-existing objects.
    ' This function should be used when script is intended to run in the session first found by the program.
    ' Param transaction - the transaction to execute updates in.
    ' Return void.
    Public Sub InitWithTransaction(ByVal transaction As String)
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

    ' Function to initialize a connection to SAP in a new GuiSession.
    ' This initialized should be used in most cases not to overwrite the users other transactions.
    ' Param transaction - the transaction to execute updates in.
    ' Return void.
    Public Sub InitWithNewSessionAndTransaction(ByVal transaction As String)
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

    ' Function to close the connection with SAP.
    ' Return void.
    Public Sub CloseConnection()
        If createdSession = True Then
            _session.FindById("wnd[0]").Close()
        End If
        ' Barber pole is not disabled. 
        _session = Nothing
        con = Nothing
        app = Nothing
        sapGuiAuto = Nothing
    End Sub

    ' Function to reset the connection with SAP, simpler form of closeConnection()
    ' Return void. 
    Private Sub resetSession()
        _session = Nothing
        sapGuiAuto = Nothing
    End Sub

    Public Sub resetTransaction(ByVal transaction As String)
        If Not _session.findById("wnd[1]", False) Is Nothing Then
            _session.findById("wnd[1]").Close()
        End If
        _session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n" & transaction
        _session.FindById("wnd[0]").sendVKey(0)
    End Sub

End Module
