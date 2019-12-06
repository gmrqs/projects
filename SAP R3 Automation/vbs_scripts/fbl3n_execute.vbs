If Not IsObject(application) Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
    Set connection = application.Children(0)
End If
If Not IsObject(session) Then
    Set session = connection.Children(0)
End If
If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

'sets maximum number of results
session.findById("wnd[0]/usr/txtPA_NMAX").text = "99999999"
session.findById("wnd[0]/usr/txtPA_NMAX").setFocus
session.findById("wnd[0]/usr/txtPA_NMAX").caretPosition = 10
'F8 - Runs query
session.findById("wnd[0]/usr/txtPA_NMAX").press