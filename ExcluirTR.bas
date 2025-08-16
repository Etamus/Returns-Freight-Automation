Attribute VB_Name = "Módulo17"
Sub EXCLUIR_TR()

Application.ScreenUpdating = False
Windows("Criação Transporte.xlsm").Activate
Sheets("Alterar RFQ e TR").Select
Range("E1").Select

    QTYLINHAS3 = Range("E10000").End(xlUp).Row
    QTYLINHAS2 = Range("F10000").End(xlUp).Row
    QTYLINHAS3 = QTYLINHAS2 + 1
    Range("E" & QTYLINHAS3).Select

Dim TR
nl = Application.WorksheetFunction.CountA(Range("E:E")) - 1
For i = 0 To nl
TR = ActiveCell.Offset(0, 0).Value
If TR = "" Then
GoTo FIM
End If

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nyt02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC").Select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1028/btnSCD_DISPLAY_1").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nyt02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press

Windows("Criação Transporte.xlsm").Activate
Sheets("Alterar RFQ e TR").Select
ActiveCell.Offset(0, 1).Value = "Transporte Excluído"
ActiveCell.Offset(1, 0).Select

Next

FIM:
On Error Resume Next
session.findById("wnd[0]").sendVKey 12
On Error GoTo 0

MsgBox "Finalizado."

Sheets("Analisar NF").Select

End Sub

