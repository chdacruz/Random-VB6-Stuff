VERSION 5.00
Begin VB.Form frmRelatorio 
   Caption         =   "Relatório Anual"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCopias 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   4440
      Width           =   945
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame frame_Escolha 
      Caption         =   "Selecione um Relatório"
      Height          =   3495
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      Begin VB.ComboBox cboMensal 
         Height          =   315
         ItemData        =   "frm_Relatorio.frx":0000
         Left            =   1560
         List            =   "frm_Relatorio.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtFim 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   2880
         Width           =   945
      End
      Begin VB.TextBox txtInicio 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   2400
         Width           =   945
      End
      Begin VB.TextBox txtAnual 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1300
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Período"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton optMensal 
         Caption         =   "Mensal"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optAnual 
         Caption         =   "Anual"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Fim:"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Início:"
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   2520
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Cópias:"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "frmRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCopias As Integer
Dim opt As String

Private Function isFieldEmpty(txtField As TextBox) As Boolean
    isFieldEmpty = False
    If txtField = "" Then
        isFieldEmpty = True
        MsgBox ("Campo Inválido ou não preenchido")
    End If
End Function

Private Function isCboEmpty(cbBox As ComboBox) As Boolean
    isCboEmpty = False
    If cbBox = "" Then
        isCboEmpty = True
        MsgBox ("Nenhum mês foi selecionado")
    End If
End Function


Private Sub cboMensal_Change()
    opt = cboMensal.Text
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    If IsNumeric(nCopias) And nCopias > 0 Then
        Select Case opt
            Case "anual"
                If isFieldEmpty(txtAnual) Then
                    Exit Sub
                End If
                
                Dim i As Integer
                For i = 1 To nCopias
                    MsgBox ("Imprimindo página " & i & " de " & nCopias & " do Relatório Anual de " & txtAnual.Text)
                Next
            Case "mensal"
                If isCboEmpty(cboMensal) Then
                    Exit Sub
                End If
                
                Dim j As Integer
                For j = 1 To nCopias
                    MsgBox ("Imprimindo página " & j & " de " & nCopias & " do Relatório do mês de " & cboMensal.List(cboMensal.ListIndex))
                Next
            Case "periodo"
                If isFieldEmpty(txtInicio) Or isFieldEmpty(txtFim) Then
                    Exit Sub
                End If
                
                Dim k As Integer
                For k = 1 To nCopias
                    MsgBox ("Imprimindo página " & k & " de " & nCopias & " do Relatório do Período de " & txtInicio & " até " & txtFim)
                Next
            End Select
    Else
        MsgBox ("Número de cópias inválido")
    End If
         
        
End Sub
Private Sub txtCopias_Change()
    If IsNumeric(txtCopias.Text) Then
        nCopias = CInt(txtCopias.Text)
    End If
    

End Sub

Private Sub Form_Load()
    cboMensal.AddItem "Janeiro"
    cboMensal.AddItem "Fevereiro"
    cboMensal.AddItem "Março"
    cboMensal.AddItem "Abril"
    cboMensal.AddItem "Maio"
    cboMensal.AddItem "Junho"
    cboMensal.AddItem "Julho"
    cboMensal.AddItem "Agosto"
    cboMensal.AddItem "Setembro"
    cboMensal.AddItem "Outubro"
    cboMensal.AddItem "Novembro"
    cboMensal.AddItem "Dezembro"
    
End Sub

Private Sub optAnual_Click()
    cboMensal.Enabled = False
    txtInicio.Enabled = False
    txtFim.Enabled = False
    txtAnual.Enabled = True
    
    opt = "anual"

End Sub

Private Sub optMensal_Click()
    txtAnual.Enabled = False
    txtInicio.Enabled = False
    txtFim.Enabled = False
    cboMensal.Enabled = True
    
    opt = "mensal"
    
End Sub

Private Sub optPeriodo_Click()
    txtAnual.Enabled = False
    cboMensal.Enabled = False
    txtInicio.Enabled = True
    txtFim.Enabled = True
    
    opt = "periodo"
    
End Sub
