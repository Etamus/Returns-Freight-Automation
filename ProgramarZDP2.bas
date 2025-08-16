Attribute VB_Name = "Módulo11"
Sub Programar_ZDP2()

Application.ScreenUpdating = False

Windows("Criação Transporte.xlsm").Activate
Sheets("Entrada").Select
Range("B11").Select
DTJOB = Range("B5").Value

QTYLINHAS = Range("B10000").End(xlUp).Row
ActiveSheet.Range("$B$11:$B$" & QTYLINHAS).RemoveDuplicates Columns:=1, Header:= _
        xlYes
QTYLINHAS = ""

Windows("Criação Transporte.xlsm").Activate
Sheets("Entrada").Select
Range("B11").Select
QTYOI = Range("B100000").End(xlUp).Row
Range("B11:B" & QTYOI).Select
Selection.Copy


If Not IsObject(app) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set app = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = app.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject app, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzdopp010"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_AUART-LOW").Text = "ZDP2"
session.findById("wnd[0]/usr/ctxtS_AUART-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = DTJOB
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").caretPosition = 10
session.findById("wnd[0]").sendVKey 9
session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "lp01"
session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 12

QTYOI = ""

Windows("Criação Transporte").Activate
Sheets("Entrada").Select
Range("G2") = "Programado"
Application.CutCopyMode = False
Range("G1").Select

frmMenu.Hide
MsgBox "ZDP2 Programado."

End Sub

