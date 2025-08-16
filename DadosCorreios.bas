Attribute VB_Name = "Módulo45"
Sub CopiarDadosCorreios()

    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaOrigem As Long
    Dim linhaDestinoA As Long
    Dim linhaDestinoE As Long
    Dim i As Long
    Dim dictE As Object
    Set dictE = CreateObject("Scripting.Dictionary")
    
    Set wsOrigem = ThisWorkbook.Sheets("Correios")
    Set wsDestino = ThisWorkbook.Sheets("Alterar RFQ e TR")
    
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "F").End(xlUp).Row
    linhaDestinoA = 2
    linhaDestinoE = 2

    For i = 2 To ultimaLinhaOrigem
        If wsOrigem.Cells(i, "F").Value <> "5002359" Then
        
            wsDestino.Cells(linhaDestinoA, "A").Value = wsOrigem.Cells(i, "D").Value
            linhaDestinoA = linhaDestinoA + 1
            
            Dim valorE As String
            valorE = Trim(wsOrigem.Cells(i, "E").Value)
            If valorE <> "" And Not dictE.exists(valorE) Then
                dictE.Add valorE, True
                wsDestino.Cells(linhaDestinoE, "E").Value = valorE
                linhaDestinoE = linhaDestinoE + 1
            End If
            
        End If
    Next i

    ' Preenche "01" na coluna B
    Dim linhaB As Long
    linhaB = 2
    Do While wsDestino.Cells(linhaB, "A").Value <> ""
        wsDestino.Cells(linhaB, "B").Value = "'01"
        linhaB = linhaB + 1
    Loop

Call Alterar_RFQ_Coletiva

End Sub
