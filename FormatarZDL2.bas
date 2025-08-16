Attribute VB_Name = "Módulo4"
Sub Formatar_ZDL2()

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
Windows("ZDL2").Activate
Sheets("ZDL2").Select

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

Sheets("ZDL2").Select

cont = 2
Do While Cells(cont, 17) <> ""

    If IsNumeric(Cells(cont, 22).Value) And IsNumeric(Cells(cont, 36).Value) Then
        If Cells(cont, 22).Value * 1 = 509 And Cells(cont, 36).Value * 1 <> 5002359 Then
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

            If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").Text = "01" Then
                GoTo rfqok
            End If

            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").Text = "01"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press

            On Error Resume Next
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

rfqok:
            On Error GoTo 0

            Do While Cells(cont, 17).Value = Cells(cont + 1, 17).Value
                cont = cont + 1
            Loop

        End If
    End If

    cont = cont + 1

Loop


   
   
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

QTYLINHA = Range("Q1000000").End(xlUp).Row
Range("V1").Select
    ActiveSheet.Range("$A$1:AW" & QTYLINHA).AutoFilter Field:=22, Criteria1:=Array( _
        "125", "025", "130", "150", "159", "160", "181", "411", "441", "508", "509", "671"), Operator:=xlFilterValues
    Rows("2:" & QTYLINHA + 1).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
QTYLINHA = ""
ActiveSheet.ShowAllData


cont = 2
Do While Cells(cont, 17) <> ""
If Cells(cont, 36).Value * 1 = 5002359 Then
Rows(cont & ":" & cont).Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.Delete Shift:=xlUp
Else
cont = cont + 1
End If
Loop

Dim ORDEM, FATURAMENTO, FORNECIMENTO
Windows("ZDL2").Activate
Sheets("ZDL2").Select
QTYLINHAS = Range("Q100000").End(xlUp).Row
Range("Q2").Select
linha = 2
FATURAMENTO = 41
FORNECIMENTO = 45
While ActiveCell.Offset(0, 0).Value <> ""
ORDEM = ActiveCell.Offset(0, 0).Value
If Cells(linha, FATURAMENTO).Value <> "" And Cells(linha, FORNECIMENTO).Value = "" Then
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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NVA02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = ORDEM
session.findById("wnd[0]").sendVKey 0
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").Key = " "
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").SetFocus
session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume Next
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
ActiveCell.Offset(0, 28).Value = "REMESSA DESBLOQUEADA"
End If
linha = linha + 1
ActiveCell.Offset(0 + 1, 0).Select
While ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(0 - 1, 0).Value
If ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(0 - 1, 0).Value Then
ActiveCell.Offset(0 + 1, 0).Select
linha = linha + 1
End If
Wend
Wend
QTYLINHAS = ""
linha = ""
FATURAMENTO = ""
FORNECIMENTO = ""
ORDEM = ""

Windows("ZDL2").Activate
Sheets("ZDL2").Select
QTYLINHAS2 = Range("Q100000").End(xlUp).Row
Range("AS2:AS" & QTYLINHAS2).Select
Selection.Copy
QTYLINHAS2 = ""

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
    Windows("ZDL2.XLS").Activate
    Sheets("ZDL2").Select
    Qtremessa = Range("Q100000").End(xlUp).Row
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
    Windows("ZDL2.XLS").Activate
    Range("AX:AX,AY:AY").Select
    Range("AY1").Activate
    Selection.NumberFormat = "m/d/yyyy"
    Windows("ZDL2.XLS").Activate
    Columns("AX:AY").Select
    Columns("AX:AAY").EntireColumn.AutoFit
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
        
    Windows("ZDL2.XLS").Activate
    Columns("AS:AS").Select
    Selection.TextToColumns Destination:=Range("AS1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("AS1").Select
    
    Windows("ZDL2.XLS").Activate
    'Range("Y1").Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
      '  SkipBlanks:=False, Transpose:=False
    'Application.CutCopyMode = False
    'Range("Z1").Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        'SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
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

