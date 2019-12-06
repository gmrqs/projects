'this block is in every vbs script it is the connection between system and SAP
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

'opens SAP
session.findBuId("wnd[0]").maximize
'types in the transcation code
session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL3N"
'presses enter
session.findById("wnd[0]").sendKey 0 