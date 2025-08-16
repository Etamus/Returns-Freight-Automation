Attribute VB_Name = "Módulo42"
Sub ANALISAR_FORMATO_NF()

    Application.ScreenUpdating = False
    ChDir "C:\temp"
    Workbooks.OpenText Filename:="C:\temp\ZDP2.xls", Origin:=xlWindows, _
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
    Windows("ZDP2").Activate
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AP1").Value = "Analisar Formato NF"
    
'elimina faturadas
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 43).Value <> "" Or Cells(cont, 47).Value <> "" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
cont = cont + 1
End If
Loop
'elimina canceladas
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 22).Value = "159" Or Cells(cont, 22).Value = "160" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
    cont = cont + 1
End If
Loop

'elimina linha sem referencia
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 41).Value = "" Or Cells(cont, 41).Value = "000000001-1" Or Cells(cont, 41).Value = "000000001-001" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
cont = cont + 1
End If
Loop

'verifica possivel erro na NFD
cont = 2
Do While Cells(cont, 17) <> ""
    If Len(Cells(cont, 41).Value) < 11 Then
    Cells(cont, 42).Value = "ANALISAR NF"
    GoTo pula
    Else
        If Left(Cells(cont, 41).Value, 9) = "000000000" Then
        Cells(cont, 42).Value = "ANALISAR NF"
        End If
    End If
pula:
cont = cont + 1
Loop


    Windows("ZDP2.XLS").Activate
    QTYLINHAS = Range("Q100000").End(xlUp).Row
    'Selection.AutoFilter
    ActiveSheet.Range("$A$1:AX" & QTYLINHAS).AutoFilter Field:=42, Criteria1:="<>"
    ActiveSheet.Range("$A$1:AX" & QTYLINHAS).AutoFilter Field:=41, Criteria1:="<>"
    Range("Q1:Q" & QTYLINHAS).Select
    'Columns("D:D").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("ZDP2").Select
    Range("AO1:AO" & QTYLINHAS).Select
    'Columns("S:S").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Planilha1").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("ZDP2").Select
    Range("AP1:AP" & QTYLINHAS).Select
    'Columns("T:T").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Planilha1").Select
    Range("C1").Select
    ActiveSheet.Paste
    Range("A1:C" & QTYLINHAS).Select
    'Columns("A:C").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:C" & QTYLINHAS).RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
      
    Windows("ZDP2.XLS").Activate
    Sheets("Planilha1").Select
    If Worksheets("Planilha1").Range("A2").Value = "" Then
    GoTo próximo1
    End If
    
    Range("A2:C2").Select
    QTYLINHAS2 = Range("A100000").End(xlUp).Row
    Range("A2:C" & QTYLINHAS2).Select
    Selection.Copy
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Analisar NF").Select
    QTYREGISTROS = Range("A1000000").End(xlUp).Row
    QTYREGISTROS = QTYREGISTROS + 1
    cont = (QTYLINHAS2 - QTYREGISTROS) + 1
    Range("A" & QTYREGISTROS).Select
    ActiveSheet.Paste
    QTYLINHAS = ""
    QTYLINHAS2 = ""
    QTYREGISTROS = ""
    
próximo1:
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
    Windows("REB").Activate
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AP1").Value = "Analisar Formato NF"
    
'elimina faturadas
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 43).Value <> "" Or Cells(cont, 47).Value <> "" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
cont = cont + 1
End If
Loop
'elimina canceladas
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 22).Value = "159" Or Cells(cont, 22).Value = "160" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
    cont = cont + 1
End If
Loop

'elimina linha sem referencia
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 41).Value = "" Or Cells(cont, 41).Value = "000000001-1" Or Cells(cont, 41).Value = "000000001-001" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
cont = cont + 1
End If
Loop

'verifica possivel erro na NFD
cont = 2
Do While Cells(cont, 17) <> ""
    If Len(Cells(cont, 41).Value) < 11 Then
    Cells(cont, 42).Value = "ANALISAR NF"
    GoTo pula1
    Else
        If Left(Cells(cont, 41).Value, 9) = "000000000" Then
        Cells(cont, 42).Value = "ANALISAR NF"
        End If
    End If
pula1:
cont = cont + 1
Loop
    
    Windows("REB.XLS").Activate
    QTYLINHAS = Range("Q100000").End(xlUp).Row
    ActiveSheet.Range("$A$1:AX" & QTYLINHAS).AutoFilter Field:=42, Criteria1:="<>"
    ActiveSheet.Range("$A$1:AX" & QTYLINHAS).AutoFilter Field:=41, Criteria1:="<>"
    Range("Q1:Q" & QTYLINHAS).Select
    'Columns("D:D").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("REB").Select
    Range("AO1:AO" & QTYLINHAS).Select
    'Columns("S:S").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Planilha1").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("REB").Select
    Range("AP1:AP" & QTYLINHAS).Select
    'Columns("T:T").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Planilha1").Select
    Range("C1").Select
    ActiveSheet.Paste
    Range("A1:C" & QTYLINHAS).Select
    'Columns("A:C").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:C" & QTYLINHAS).RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
      
    Windows("REB.XLS").Activate
    Sheets("Planilha1").Select
    If Worksheets("Planilha1").Range("A2").Value = "" Then
    GoTo próximo2
    End If
    
    Range("A2:C2").Select
    QTYLINHAS2 = Range("A100000").End(xlUp).Row
    Range("A2:C" & QTYLINHAS2).Select
    Selection.Copy
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Analisar NF").Select
    QTYREGISTROS = Range("A1000000").End(xlUp).Row
    QTYREGISTROS = QTYREGISTROS + 1
    cont = (QTYLINHAS2 - QTYREGISTROS) + 1
    Range("A" & QTYREGISTROS).Select
    ActiveSheet.Paste
    QTYLINHAS = ""
    QTYLINHAS2 = ""
    QTYREGISTROS = ""
    
próximo2:
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
    Windows("ZDL2").Activate
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AP1").Value = "Analisar Formato NF"
    
'elimina faturadas
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 43).Value <> "" Or Cells(cont, 47).Value <> "" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
cont = cont + 1
End If
Loop
'elimina canceladas
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 22).Value = "159" Or Cells(cont, 22).Value = "160" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
    cont = cont + 1
End If
Loop

'elimina linha sem referencia
cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 41).Value = "" Or Cells(cont, 41).Value = "000000001-1" Or Cells(cont, 41).Value = "000000001-001" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete
Else
cont = cont + 1
End If
Loop

'verifica possivel erro na NFD
cont = 2
Do While Cells(cont, 17) <> ""
    If Len(Cells(cont, 41).Value) < 11 Then
    Cells(cont, 42).Value = "ANALISAR NF"
    GoTo pula2
    Else
        If Left(Cells(cont, 41).Value, 9) = "000000000" Then
        Cells(cont, 42).Value = "ANALISAR NF"
        End If
    End If
pula2:
cont = cont + 1
Loop

    Windows("ZDL2.XLS").Activate
    QTYLINHAS = Range("Q100000").End(xlUp).Row
    ActiveSheet.Range("$A$1:AX" & QTYLINHAS).AutoFilter Field:=42, Criteria1:="<>"
    ActiveSheet.Range("$A$1:AX" & QTYLINHAS).AutoFilter Field:=41, Criteria1:="<>"
    Range("Q1:Q" & QTYLINHAS).Select
    'Columns("D:D").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("ZDL2").Select
    Range("AO1:AO" & QTYLINHAS).Select
    'Columns("S:S").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Planilha1").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("ZDL2").Select
    Range("AP1:AP" & QTYLINHAS).Select
    'Columns("T:T").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Planilha1").Select
    Range("C1").Select
    ActiveSheet.Paste
    Range("A1:C" & QTYLINHAS).Select
    'Columns("A:C").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:C" & QTYLINHAS).RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
      
    
    Windows("ZDL2.XLS").Activate
    Sheets("Planilha1").Select
    If Worksheets("Planilha1").Range("A2").Value = "" Then
    GoTo próximo3
    End If

    Range("A2:C2").Select
    QTYLINHAS2 = Range("A100000").End(xlUp).Row
    Range("A2:C" & QTYLINHAS2).Select
    Selection.Copy
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Analisar NF").Select
    QTYREGISTROS = Range("A1000000").End(xlUp).Row
    QTYREGISTROS = QTYREGISTROS + 1
    cont = (QTYLINHAS2 - QTYREGISTROS) + 1
    Range("A" & QTYREGISTROS).Select
    ActiveSheet.Paste
    QTYLINHAS = ""
    QTYLINHAS2 = ""
    QTYREGISTROS = ""
próximo3:

    Application.DisplayAlerts = False
    Windows("ZDP2.XLS").Activate
    ActiveWindow.Close
    Windows("REB.XLS").Activate
    ActiveWindow.Close
    Windows("ZDL2.XLS").Activate
    ActiveWindow.Close
    Application.DisplayAlerts = True
    
    Windows("Criação Transporte.xlsm").Activate
    Sheets("Analisar NF").Select
    
    Dim txtA As String
    Dim txtB As String
    
    Range("D2").Select
    Do While ActiveCell <> ""
    varQtde = Len(ActiveCell.Offset(0, 15).Value)
    If varQtde < 11 Then
    ActiveCell.Offset(0, 16).Value = "Analisar NF"
    End If
    ActiveCell.Offset(1, 0).Select
    varQtde = ""
    Loop
    
    txtA = ","
    txtB = ""
    
    Range("A1").Select

End Sub
