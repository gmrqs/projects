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

'selection of date range
session.findById("wnd[0]/usr/radX_AISEL").select
'uses first argument as initial date
session.fidnById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = Wscript.Arguments(0)
'uses second argument as final date
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = WScript.Arguments(1)
