Attribute VB_Name = "Módulo3"
Sub Formatar_REB()
Attribute Formatar_REB.VB_ProcData.VB_Invoke_Func = " \n14"


Application.ScreenUpdating = False
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
Sheets("REB").Select

    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
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

QTYLINHA = Range("Q1000000").End(xlUp).Row
Range("J1").Select
ActiveSheet.Range("$A$1:AW" & QTYLINHA).AutoFilter Field:=10, Criteria1:="<>"
    Rows("2:" & QTYLINHA + 1).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
QTYLINHA = ""
ActiveSheet.ShowAllData

cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 46).Value <> "" Then
Rows(cont & ":" & cont).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.Delete Shift:=xlUp
Else
cont = cont + 1
End If
Loop
'Correção correios
cont = 2
Do While Cells(cont, 17) <> ""
    If Cells(cont, 22).Value = "509" And Cells(cont, 36).Value <> "5002359" Then
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
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = Cells(cont, 17).Value
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
        session.findById("wnd[0]").sendVKey 0
        On Error Resume Next
        session.findById("wnd[1]").sendVKey 0
        On Error GoTo 0
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").Text = "01"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        On Error Resume Next
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        'On Error GoTo 0
        
        session.findById("wnd[0]/tbar[1]/btn[17]").press
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          5", "&Hierarchy"
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          5", "&Hierarchy"
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[0]/mbar/menu[0]/menu[4]").Select
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
        
        Dim CDC As Variant
        linha = 0
        CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
        
        Do Until CDC = ""
        If CDC = "SP Transportador" Then
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,1]").Text = "5002359"
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,1]").SetFocus
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,1]").caretPosition = 7
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[0]").sendVKey 3
        Exit Do
        Else
        linha = linha + 1
        CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
        End If
        Loop
        
igual:
        If Cells(cont, 17).Value = Cells(cont + 1, 17).Value Then
        cont = cont + 1
        GoTo igual
        End If

    End If
            cont = cont + 1
Loop

cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 36).Value = "5002359" Then
Rows(cont & ":" & cont).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.Delete Shift:=xlUp
Else
cont = cont + 1
End If
Loop

QTYLINHA = Range("Q1000000").End(xlUp).Row
Range("V1").Select
    ActiveSheet.Range("$A$1:AW" & QTYLINHA).AutoFilter Field:=22, Criteria1:=Array( _
        "025", "125", "130", "159", "160", "181", "441", "411", "508", "509", "671"), Operator:=xlFilterValues
    Rows("2:" & QTYLINHA + 1).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
QTYLINHA = ""
ActiveSheet.ShowAllData

    
    Sheets("REB").Select
    Range("AJ1").Select
    Do Until ActiveCell = ""
    ActiveCell.Offset(1, 0).Select
    If ActiveCell = "5002359" Then
    LinhaEliminar = Range("Q100000").End(xlUp).Row
    Range("AJ2").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:AW" & LinhaEliminar).AutoFilter Field:=36, Criteria1:=Array( _
    "5002359")
    Rows("2:" & LinhaEliminar + 1).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    LinhaEliminar = ""
    ActiveSheet.ShowAllData
    Exit Do
    End If
    Loop
    


Windows("REB").Activate
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
    Windows("REB.XLS").Activate
    Sheets("REB").Select
    
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
    'Windows("DtRemessa.XLSX").Activate
    Windows("REB.XLS").Activate
    Range("AX:AX,AY:AY").Select
    Range("AY1").Activate
    Selection.NumberFormat = "m/d/yyyy"
    Windows("REB.XLS").Activate
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
        
    Windows("REB.XLSX").Activate
    Columns("AS:AS").Select
    Selection.TextToColumns Destination:=Range("AS1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("AS1").Select
    
    Windows("REB.XLS").Activate
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

End Sub

