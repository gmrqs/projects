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

'with all desired companies in your clipboard(ctrl+c)

'access multi-selection in account parameter
session.findById("wnd[0]/usr/btn%_SD_BUKRS_%_APP_%-VALU_PUSH").press
'clears the list of potential undesired companies prior in the list
session.findById("wnd[1]/tbar[0]/btn[16]").press
'past the company list from clipboard to SAP
session.findById("wnd[1]/tbar[0]/btn[24]").press
'confirms selection
session.findById("wnd[1]/tbar[0]/btn[8]").press
