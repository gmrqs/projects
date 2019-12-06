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

session.findById("wnd[0]").maximize
'select option to extract table to a local file (.xls)
session.findById("wnd[0]/mbar/menu[0]/menu[3]").select
session.findById("wnd[1]/usr/sub/SUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/sub/SUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
'inputs the file path
session.findById("wnd[1]/usr/ctxtDY_PATH").text = Wscript.Arguments(0)
'inputs the file name
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = Wscript.Arguments(1)
session.findById("wnd[1]/tbar[0]/btn[11]").press