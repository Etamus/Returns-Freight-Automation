Attribute VB_Name = "Módulo1"
Sub Extrair_Bases_ZV62N()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

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

Windows("Criação Transporte").Activate
Sheets("Entrada").Select
TP_Ordem = Range("A2" & Inicio).Text
Data_Inicial = Range("B2").Text
Data_Final = Range("C2").Text
Status_Ordem = Range("D2").Text
Limpar_Selecao = Range("E2").Text
Status = Range("F2").Text

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZV62N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_AUART-LOW").Text = TP_Ordem
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = Data_Inicial
session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = Data_Final
session.findById("wnd[0]/usr/ctxtS_GBSTK-LOW").Text = Status_Ordem
session.findById("wnd[0]/usr/ctxtS_GBSTK-HIGH").Text = Status_Ordem
session.findById("wnd[0]").sendVKey 8

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "CODOC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "CODOC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").deselectColumn "DESCRICAO"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL").Select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-ILOW_I[1,0]").Text = "600"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").Text = "799"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").SetFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 3
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZDP2.xls"


On Error Resume Next
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

Windows("Criação Transporte").Activate
Sheets("Entrada").Select
Range("F2") = "Realizado"

Windows("Criação Transporte").Activate
Sheets("Entrada").Select
TP_Ordem = Range("A3").Text
Data_Inicial = Range("B3").Text
Data_Final = Range("C3").Text
Status_Ordem = Range("D3").Text
Limpar_Selecao = Range("E3").Text
Status = Range("F3").Text

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZV62N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_AUART-LOW").Text = TP_Ordem
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = Data_Inicial
session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = Data_Final
session.findById("wnd[0]/usr/ctxtS_GBSTK-LOW").Text = Status_Ordem
session.findById("wnd[0]/usr/ctxtS_GBSTK-HIGH").Text = Status_Ordem
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "CODOC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "CODOC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").deselectColumn "DESCRICAO"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL").Select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-ILOW_I[1,0]").Text = "600"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").Text = "799"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").SetFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 3
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "REB.xls"
On Error Resume Next
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press


Windows("Criação Transporte").Activate
Sheets("Entrada").Select
Range("F3") = "Realizado"

Windows("Criação Transporte").Activate
Sheets("Entrada").Select
TP_Ordem = Range("A4").Text
Data_Inicial = Range("B4").Text
Data_Final = Range("C4").Text
Status_Ordem = Range("D4").Text
Limpar_Selecao = Range("E4").Text
Status = Range("F4").Text

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZV62N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_AUART-LOW").Text = TP_Ordem
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = Data_Inicial
session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").Text = Data_Final
session.findById("wnd[0]/usr/ctxtS_GBSTK-LOW").Text = Status_Ordem
session.findById("wnd[0]/usr/ctxtS_GBSTK-HIGH").Text = Status_Ordem
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "CODOC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "CODOC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").deselectColumn "DESCRICAO"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL").Select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-ILOW_I[1,0]").Text = "600"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").Text = "799"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").SetFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/txtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 3
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZDL2.xls"

On Error Resume Next
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

Windows("Criação Transporte").Activate
Sheets("Entrada").Select
Range("F4") = "Realizado"

Dim wb As Workbook
Dim arquivos As Variant
Dim nomeArquivo As Variant

arquivos = Array("C:\temp\ZDL2.xls", "C:\temp\ZDP2.xls", "C:\temp\REB.xls")

For Each nomeArquivo In arquivos
    Set wb = Workbooks.Open(nomeArquivo)
    With wb.Sheets(1)
        .Rows("2:3").Delete
    End With
    wb.Save
    wb.Close
Next nomeArquivo

Application.DisplayAlerts = True
Application.ScreenUpdating = True

MsgBox "Finalizado."

End Sub
