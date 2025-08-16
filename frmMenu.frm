VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenu 
   ClientHeight    =   4728
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5136
   OleObjectBlob   =   "frmMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Preenche a ComboBox de Formatar
    With cmbFormatar
        .AddItem "ZDP2"
        .AddItem "REB"
        .AddItem "ZDL2"
    End With
    
    ' Preenche a ListBox de Programar Transporte
    With lstProgramar
        .AddItem "ZDP2"
        .AddItem "REB"
        .AddItem "ZDL2"
        .MultiSelect = fmMultiSelectMulti
    End With
End Sub

Private Sub btnFormatar_Click()
    Select Case cmbFormatar.Value
        Case "ZDP2"
            Call Formatar_ZDP2
        Case "REB"
            Call Formatar_REB
        Case "ZDL2"
            Call Formatar_ZDL2
        Case Else
            MsgBox "Selecione uma opção para Formatar.", vbExclamation
    End Select
End Sub

Private Sub btnProgramar_Click()
    Dim i As Integer, encontrou As Boolean
    encontrou = False
    
    For i = 0 To lstProgramar.ListCount - 1
        If lstProgramar.Selected(i) Then
            encontrou = True
            Select Case lstProgramar.List(i)
                Case "ZDP2"
                    Call Programar_ZDP2
                Case "REB"
                    Call Programar_REB
                Case "ZDL2"
                    Call Programar_ZDL2
            End Select
        End If
    Next i
    
    If Not encontrou Then
        MsgBox "Selecione pelo menos uma opção para Programar.", vbExclamation
    End If
End Sub

Private Sub btnFechar_Click()
    Unload Me
End Sub
