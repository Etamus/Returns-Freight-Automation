Attribute VB_Name = "Módulo20"
Sub LANÇAMENTO_01_ZDL2_REB()
Attribute LANÇAMENTO_01_ZDL2_REB.VB_ProcData.VB_Invoke_Func = " \n14"

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
    Windows("ZDL2.xls").Activate
    Sheets("ZDL2").Select
    
        Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    Columns("S:S").Select
    Selection.Cut
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight
    Range("AJ1").Select
    Do Until ActiveCell = "5002359"
    ActiveCell.Offset(1, 0).Select
    If ActiveCell = "5002359" Then
    QTYLINHA = Range("Q1000000").End(xlUp).Row
    ActiveSheet.Range("$A$1:$AW" & QTYLINHA).AutoFilter Field:=36, Criteria1:= _
        "5002359"
    ActiveSheet.Range("$A$1:$AW" & QTYLINHA).AutoFilter Field:=22, Criteria1:=Array( _
        "181", "508", "509"), Operator:=xlFilterValues
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=46, Criteria1:="="
    Range("Q1:R" & QTYLINHA).Select
    'Columns("D:E").Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Exit Do
    End If
    Loop
    QTYLINHA = ""
    QTYLINHA = Range("A1000000").End(xlUp).Row
    Range("A2:B" & QTYLINHA).Select
    Selection.Copy
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Correios").Select
    
    Range("A1").Select
    QTYLINHAS3 = Range("A100000").End(xlUp).Row
    QTYLINHAS2 = Range("A100000").End(xlUp).Row
    QTYLINHAS3 = QTYLINHAS2 + 1
    Range("A" & QTYLINHAS3).Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("ENTRADA").Select
    QTYLINHA = ""
    QTYLINHAS2 = ""
    QTYLINHAS3 = ""
    
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
    Windows("REB.xls").Activate
    Sheets("REB").Select
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    Columns("S:S").Select
    Selection.Cut
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight
    Range("AJ1").Select
    Do Until ActiveCell = "5002359"
    ActiveCell.Offset(1, 0).Select
    If ActiveCell = "5002359" Then
    QTYLINHA = Range("Q1000000").End(xlUp).Row
    ActiveSheet.Range("$A$1:$AW" & QTYLINHA).AutoFilter Field:=36, Criteria1:= _
        "5002359"
    ActiveSheet.Range("$A$1:$AW" & QTYLINHA).AutoFilter Field:=22, Criteria1:=Array( _
        "181", "508", "509"), Operator:=xlFilterValues
    ActiveSheet.Range("$A$1:$AW$" & QTYLINHA).AutoFilter Field:=46, Criteria1:="="
    Range("Q1:R" & QTYLINHA).Select
    'Columns("D:E").Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Exit Do
    End If
    Loop

    QTYLINHA = Range("A1000000").End(xlUp).Row
    Range("A2:B" & QTYLINHA).Select
    Selection.Copy
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Correios").Select
    
    Range("A1").Select
    QTYLINHAS3 = Range("A100000").End(xlUp).Row
    QTYLINHAS2 = Range("A100000").End(xlUp).Row
    QTYLINHAS3 = QTYLINHAS2 + 1
    Range("A" & QTYLINHAS3).Select
    ActiveSheet.Paste
    Windows("Criação Transporte.xlsm").Activate
    Range("A1").Select
    Sheets("Correios").Select
    
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:=".", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    QTYLINHA = ""
    
    Application.DisplayAlerts = False
    Windows("ZDL2.XLS").Activate
    ActiveWindow.Close
    Windows("REB.XLS").Activate
    ActiveWindow.Close
    Application.DisplayAlerts = True
    
    MsgBox "CONCLUÍDO, ANALIZAR DATAS, VERIFICAR SE TEM TRANSPORTE CRIADO E LANÇAR 01"

End Sub
