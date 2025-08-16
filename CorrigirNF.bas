Attribute VB_Name = "Módulo43"
Sub CORRIGIR_FORMATO_NF()

    Application.ScreenUpdating = False
    
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Analisar NF").Select
    Range("A2").Select
    linha = Range("A1000000").End(xlUp).Row
    LinhaSub = Range("E1000000").End(xlUp).Row
    linha = LinhaSub + 1
    Range("A" & linha).Select
    
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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0

    Do While ActiveCell <> ""
    ORDEM = ActiveCell.Offset(0, 0).Value
    NUMERO_NF = ActiveCell.Offset(0, 3).Value
    If ActiveCell <> "" Then
    
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = ORDEM
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
On Error Resume Next
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = "e" & NUMERO_NF
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").caretPosition = 12
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press

    Windows("Criação Transporte.xlsm").Activate
    Sheets("Analisar NF").Select
    ActiveCell.Offset(0, 4).Value = "Corrigido."
    End If
    ActiveCell.Offset(1, 0).Select
    ORDEM = ""
    NUMERO_NF = ""
    
    Loop
    
    MsgBox "Corrigido."

    Sheets("Entrada").Select

End Sub

