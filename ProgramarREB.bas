Attribute VB_Name = "Módulo12"
Sub Programar_REB()

Application.ScreenUpdating = False

Windows("Criação Transporte.xlsm").Activate
Sheets("Entrada").Select
Range("C11").Select
DTJOB = Range("B5").Value

QTYLINHAS = Range("C10000").End(xlUp).Row
ActiveSheet.Range("$C$11:$D$" & QTYLINHAS).RemoveDuplicates Columns:=1, Header:= _
        xlYes
QTYLINHAS = ""

    QTYLINHAS2 = Range("C10000").End(xlUp).Row
    Range("C10:D10").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Entrada").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Entrada").Sort.SortFields.Add Key:=Range( _
        "D16:D" & QTYLINHAS2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Entrada").Sort
        .SetRange Range("D10:D" & QTYLINHAS2)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    QTYLINHAS2 = ""

Windows("Criação Transporte.xlsm").Activate
Sheets("Entrada").Select
Range("C11").Select
QTYOI = Range("C100000").End(xlUp).Row
Range("C11:C" & QTYOI).Select
QTYAT = Range("D100000").End(xlUp).Row
Range("D11:D" & QTYAT).Select

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

Range("D17:D" & QTYAT).Select
Selection.Copy

session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press

Range("C17:C" & QTYOI).Select
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
Range("G3") = "Programado"
Application.CutCopyMode = False
Range("G1").Select

frmMenu.Hide
MsgBox "REB Programado."

End Sub



