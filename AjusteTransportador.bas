Attribute VB_Name = "Módulo44"
Sub AJUSTE_TRANSPORTADOR_CORREIOS()

Dim Teste As Variant

    Application.ScreenUpdating = False
    ChDir "C:\temp"
    Workbooks.OpenText Filename:="C:\temp\ZDL2.xls", Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 4), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 4), Array(8, 1), Array(9, 4), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 4), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 4), Array(21, 1), Array(22, 1), _
        Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array( _
        29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), _
        Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array( _
        42, 1), Array(43, 1), Array(44, 4), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 4), _
        Array(49, 1), Array(50, 1)), TrailingMinusNumbers:=True
    Columns("C:C").EntireColumn.AutoFit
    Windows("ZDL2.XLS").Activate
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    Range("V1").Select
    Do Until ActiveCell = "509"
    Teste = ActiveCell.Offset(1, -5).Value
    If ActiveCell.Offset(1, -5).Value = "" Then
    GoTo pula1
    End If
    ActiveCell.Offset(1, 0).Select
    If ActiveCell = "509" Then
    QTYLINHA = Range("Q1000000").End(xlUp).Row
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=22, Criteria1:="509"
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=46, Criteria1:="="
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=10, Criteria1:="="
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=41, Criteria1:="<>"
    Range("Q1:Q" & QTYLINHA).Select
    'Columns("D:D").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("ZDL2").Select
    Range("A1:A" & QTYLINHA).Select
    'Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    Worksheets(2).Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("ZDL2").Select
    Range("AJ1:AK" & QTYLINHA).Select
    'Columns("Q:R").Select
    Application.CutCopyMode = False
    Selection.Copy
    Worksheets(2).Select
    Range("C1").Select
    ActiveSheet.Paste
    QTYLINHA2 = Range("A1000000").End(xlUp).Row
    Range("A2:D2").Select
    Range("A2:D" & QTYLINHA2).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Correios").Select
    Range("D1").Select
    QTYLINHAS3 = Range("D100000").End(xlUp).Row
    QTYLINHAS4 = Range("D100000").End(xlUp).Row
    QTYLINHAS3 = QTYLINHAS4 + 1
    Range("D" & QTYLINHAS3).Select
    ActiveSheet.Paste
    Range("D1").Select
    Sheets("ENTRADA").Select
    QTYLINHA = ""
    QTYLINHAS2 = ""
    QTYLINHAS3 = ""
    QTYLINHAS4 = ""
    Exit Do
    End If
    Loop
pula1:
    ChDir "C:\temp"
    Workbooks.OpenText Filename:="C:\temp\REB.xls", Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 4), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 4), Array(8, 1), Array(9, 4), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 4), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 4), Array(21, 1), Array(22, 1), _
        Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array( _
        29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), _
        Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array( _
        42, 1), Array(43, 1), Array(44, 4), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 4), _
        Array(49, 1), Array(50, 1)), TrailingMinusNumbers:=True
    Columns("C:C").EntireColumn.AutoFit
    Windows("REB.XLS").Activate
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    Range("V1").Select
    Do Until ActiveCell = "509"
    Teste = ActiveCell.Offset(1, -5).Value
    If ActiveCell.Offset(1, -5).Value = "" Then
    GoTo pula2
    End If
    ActiveCell.Offset(1, 0).Select
    If ActiveCell = "509" Then
    QTYLINHA = Range("Q1000000").End(xlUp).Row
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=22, Criteria1:="509"
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=46, Criteria1:="="
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=10, Criteria1:="="
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=41, Criteria1:="<>"
    Range("Q1:Q" & QTYLINHA).Select
    'Columns("D:D").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("REB").Select
    Range("A1:A" & QTYLINHA).Select
    'Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    Worksheets(2).Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("REB").Select
    Range("AJ1:AK" & QTYLINHA).Select
    'Columns("Q:R").Select
    Application.CutCopyMode = False
    Selection.Copy
    Worksheets(2).Select
    Range("C1").Select
    ActiveSheet.Paste
    QTYLINHA2 = Range("A1000000").End(xlUp).Row
    Range("A2:D2").Select
    Range("A2:D" & QTYLINHA2).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Correios").Select
    Range("D1").Select
    QTYLINHAS3 = Range("D100000").End(xlUp).Row
    QTYLINHAS4 = Range("D100000").End(xlUp).Row
    QTYLINHAS3 = QTYLINHAS4 + 1
    Range("D" & QTYLINHAS3).Select
    ActiveSheet.Paste
    Range("D1").Select
    Sheets("ENTRADA").Select
    QTYLINHA = ""
    QTYLINHAS2 = ""
    QTYLINHAS3 = ""
    QTYLINHAS4 = ""
    Exit Do
    End If
    Loop
pula2:
    Application.DisplayAlerts = False
    Windows("ZDL2.XLS").Activate
    ActiveWindow.Close
    Windows("REB.XLS").Activate
    ActiveWindow.Close
    Sheets("ENTRADA").Select
    Application.DisplayAlerts = True
    
End Sub


