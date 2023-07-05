VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public connection, session As Object
Private Sub Class_Initialize()
create_session
End Sub
Public Function create_session()

On Error GoTo SAP_Anwendung
If Not IsObject(SAPApplication) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SAPApplication = SapGuiAuto.GetScriptingEngine
End If

On Error GoTo SAP_Anmeldung
If Not IsObject(connection) Then
   Set connection = SAPApplication.Children(0)
End If

If Not IsObject(session) Then
   Set session = connection.Children(0)
End If

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

For Each ses In connection.Children()
    If ses.Info.transaction = "SESSION_MANAGER" Then
       Set session = ses
       Exit Function
    End If
Next

If connection.Children.Count < 6 Then
    Dim current_ses As Object
    
    Set current_ses = connection.Children(connection.sessions.Count - 1) 'Current session
    Set session = current_ses
    
    session.createsession
    While session Is current_ses
        Application.Wait (Now + TimeValue("00:00:01"))
        Set session = connection.Children(connection.sessions.Count - 1) 'New session
    Wend
Else
    MsgBox "Zuviele Modi in SAP geöffnet. Bitte einen Modus schließen!"
End If

Exit Function


SAP_Anwendung:
    MsgBox "Bitte SAPGUI starten und anmelden!"
SAP_Anmeldung:
    MsgBox "Bitte in SAP anmelden!"
End Function





