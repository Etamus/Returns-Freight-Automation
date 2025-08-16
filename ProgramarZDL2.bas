Attribute VB_Name = "Módulo13"
Sub Programar_ZDL2()

Application.ScreenUpdating = False

Windows("Criação Transporte.xlsm").Activate
Sheets("Entrada").Select
Range("E11").Select
DTJOB = Range("B5").Value
      
QTYLINHAS = Range("E10000").End(xlUp).Row
ActiveSheet.Range("$E$11:$F$" & QTYLINHAS).RemoveDuplicates Columns:=1, Header:= _
        xlYes
QTYLINHAS = ""

    QTYLINHAS2 = Range("E10000").End(xlUp).Row
    Range("E10:F10").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Entrada").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Entrada").Sort.SortFields.Add Key:=Range( _
        "F10:F" & QTYLINHAS2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Entrada").Sort
        .SetRange Range("F10:F" & QTYLINHAS2)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    QTYLINHAS2 = ""

Windows("Criação Transporte.xlsm").Activate
Sheets("Entrada").Select
Range("E11").Select
QTYOI = Range("E100000").End(xlUp).Row
Range("E11:E" & QTYOI).Select
QTYAT = Range("F100000").End(xlUp).Row
Range("F11:F" & QTYAT).Select

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
Range("G4") = "Programado"
Application.CutCopyMode = False
Range("G1").Select

frmMenu.Hide
MsgBox "ZDL2 Programado."

End Sub


