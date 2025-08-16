Attribute VB_Name = "Módulo52"
Sub Gerar_TR_Cliente_ZDL2()

    Windows("Criação Transporte.xlsm").Activate
    Sheets("Gerar TR por cliente").Select
    num = Application.WorksheetFunction.CountA(Range("D:D"))
    con = Application.WorksheetFunction.CountA(Range("D:D"))
    
    Do Until con = 1
    
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Gerar TR por cliente").Select
    Cliente = Range("D" & con)
    
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Gerar TR por cliente").Select
    Range("A1").Select
    
    nl = Application.WorksheetFunction.CountA(Range("A:A"))

    ActiveSheet.Range("$A$1:$B$" & nl).AutoFilter Field:=2, Criteria1:= _
        Cliente
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Entrada").Select
    Range("E16").Select
    ActiveSheet.Paste

Windows("Criação Transporte.xlsm").Activate
Sheets("Entrada").Select
Range("E17").Select
QTYOI = Range("E100000").End(xlUp).Row
Range("E17:E" & QTYOI).Select
QTYAT = Range("F100000").End(xlUp).Row
Range("F17:F" & QTYAT).Select
DTJOB = Range("B10").Value

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr080"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_S_KUNNR_%_APP_%-VALU_PUSH").press

Range("F17:F" & QTYAT).Select
Selection.Copy

session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press

Range("E17:E" & QTYOI).Select
Selection.Copy

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
QTYAT = ""

Windows("Criação Transporte").Activate
Sheets("Entrada").Select
Range("G8") = "Programado"
Application.CutCopyMode = False
Range("G5").Select

    Windows("Criação Transporte").Activate
    Sheets("Entrada").Select
    Range("E17:F17").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("E16").Select
    
    con = con - 1
    
    If con = 1 Then
    Sheets("Gerar TR por cliente").Select
    ActiveSheet.ShowAllData
    Range("A1").Select
    GoTo FIM
    End If
    
    Loop
      
FIM:
MsgBox "ENCERRADO", vbInformation

End Sub

