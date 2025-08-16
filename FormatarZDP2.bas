Attribute VB_Name = "Módulo2"
Sub Formatar_ZDP2()
Attribute Formatar_ZDP2.VB_ProcData.VB_Invoke_Func = " \n14"

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
Sheets("ZDP2").Select

    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    
Rows("1:1").Select
Selection.AutoFilter
QTYLINHA = Range("Q1000000").End(xlUp).Row
ActiveSheet.Range("$A$1:AW" & QTYLINHA).AutoFilter Field:=1, Criteria1:="<>"
Rows("2:" & QTYLINHA + 1).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.Delete Shift:=xlUp
QTYLINHA = ""
ActiveSheet.ShowAllData


cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 22).Value * 1 = 159 Or Cells(cont, 22).Value * 1 = 160 Or Cells(cont, 22).Value * 1 = 671 Then
Rows(cont & ":" & cont).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.Delete Shift:=xlUp
Else
cont = cont + 1
End If
Loop

cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 10).Value <> "" Then
Rows(cont & ":" & cont).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.Delete Shift:=xlUp
Else
cont = cont + 1
End If
Loop

Sheets("ZDP2").Select
Range("AT1").Select

cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 43).Value <> "" Or Cells(cont, 47).Value <> "" Then
Rows(cont & ":" & cont).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.Delete Shift:=xlUp
Else
cont = cont + 1
End If
Loop

Windows("ZDP2.XLS").Activate
Sheets("ZDP2").Select
If Worksheets("ZDP2").Range("Q2").Value = "" Then
MsgBox "NÃO HÁ INPUT AGUARDANDO TRANSPORTE PARA TIPO ORDEM ZDP2"
Application.DisplayAlerts = False
Windows("ZDP2.XLS").Activate
ActiveWindow.Close
Application.DisplayAlerts = True
GoTo FIM
End If

Range("A1").Select

    Windows("ZDP2").Activate
    Sheets("ZDP2").Select
    
'Cancelar MT e MS para Jlle
cont = 2
Do While Cells(cont, 17) <> ""
    If Cells(cont, 27).Value = "MT" Or Cells(cont, 27).Value = "MS" Then
        If Cells(cont, 33).Value = "1109" Then
        canc = Cells(cont, 17).Value
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
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = canc
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
        session.findById("wnd[0]").sendVKey 0
        On Error Resume Next
        session.findById("wnd[1]").sendVKey 0
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").Key = "160"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").SetFocus
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07").Select
        session.findById("wnd[0]/tbar[1]/btn[34]").press
        session.findById("wnd[1]/usr/cmbRV45A-S_ABGRU").Key = "60"
        session.findById("wnd[1]").sendVKey 0
        session.findById("wnd[1]/tbar[0]/btn[7]").press
volta:
        On Error Resume Next
        txtd = session.findById("wnd[2]/usr/txtMESSTXT1").Text
        If txtd <> "" Then
        session.findById("wnd[0]").sendVKey 0
        txtd = ""
        GoTo volta
        On Error GoTo 0
        End If
        
        'On Error Resume Next
        'barra = session.findById("wnd[0]/sbar").Text
        'If barra = "Não foi efetuada qualquer modificação de dados" Then
        'On Error GoTo 0
        'GoTo texto
        'End If
        'On Error Resume Next
        
'texto:
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
            
            If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = "" Then
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = "e1-1"
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").SetFocus
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").caretPosition = 3
            End If
            

   

        
        ' texto
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0005", "Column1"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0005", "Column1"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0005", "Column1"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = "Conforme definição DOPP, MT e MS não retorna para 1109"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 34, 34
        
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        
        On Error Resume Next
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        On Error GoTo 0
        
        'On Error Resume Next
        'barra = session.findById("wnd[0]/sbar").Text
            'If barra = "Não foi efetuada qualquer modificação de dados" Then
            'On Error GoTo 0
            'GoTo texto5
            'End If
        'On Error Resume Next

        Windows("ZDP2").Activate
        Sheets("ZDP2").Select
cancel:
        Rows(cont & ":" & cont).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Delete Shift:=xlUp
            If Cells(cont, 17).Value = canc Then
            GoTo cancel
            End If
        Else
        cont = cont + 1
        End If
    
    Else
    cont = cont + 1
    End If
Loop
    
Windows("ZDP2").Activate
QTYLINHAS = Range("Q100000").End(xlUp).Row
Range("AS2:AS" & QTYLINHAS).Select
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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzbse16"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "likp"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[1]").Select
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Temp"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "DtRemessa.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
On Error Resume Next
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]").sendVKey 12

    Workbooks.Open Filename:="C:\Temp\DtRemessa.XLSX"
    Windows("ZDP2.XLS").Activate
    Sheets("ZDP2").Select
    
    Qtremessa = Range("Q10000").End(xlUp).Row
    Range("AS2:AS" & Qtremessa).Select
       
    Range("AX1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Data Criação"
    Range("AY1").Select
    ActiveCell.FormulaR1C1 = "Data trabalho"
    Range("AX2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-5],[DtRemessa.XLSX]Sheet1!C1:C2,2,0),""DESCONSIDERAR"")"
    Range("AY2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(WORKDAY(RC[-1],3),""DESCONSIDERAR"")"
    Range("AX2:AY2").Select
    Selection.Copy
    Range("AX2:AY2" & Qtremessa).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("AX1").Select
    Windows("DtRemessa.XLSX").Activate
    Windows("ZDP2.XLS").Activate
    Range("AX:AX,AY:AY").Select
    Range("AY1").Activate
    Selection.NumberFormat = "m/d/yyyy"
    Windows("ZDP2.XLS").Activate
    Columns("AX:AY").Select
    Columns("AX:AY").EntireColumn.AutoFit
    Rows("1:1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("AY1").Select
    Windows("Criação Transporte.xlsm").Activate
    
    Windows("DtRemessa.XLSX").Activate
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
    Windows("ZDP2.XLS").Activate
    Columns("AS:AS").Select
    Selection.TextToColumns Destination:=Range("AS1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("AS1").Select
    
    Windows("ZDP2.XLS").Activate
    Columns("AX:AY").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("AY1").Select
    
    Workbooks("DtRemessa.XLSX").Close SaveChanges:=False
    Windows("Criação Transporte.xlsm").Activate
    MsgBox "Extração Concluída."
    frmMenu.Hide
FIM:
    
End Sub
