VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   5715
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculadora 
      Caption         =   "Calculadora"
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdRelatorio 
      Caption         =   "Relatório Anual"
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdLista 
      Caption         =   "Lista"
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuSalvar 
         Caption         =   "Salvar"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "Editar"
      Begin VB.Menu mnuBackground 
         Caption         =   "Background"
      End
      Begin VB.Menu mnuFonte 
         Caption         =   "Fonte"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculadora_Click()
    frmCalculadora.Show
End Sub

Private Sub cmdLista_Click()
    frmLista.Show

End Sub

Private Sub cmdRelatorio_Click()
    frmRelatorio.Show
    
End Sub

Private Sub mnuBackground_Click()
    'frmLista.BackColor = Color.Green
    cmdLista.BackColor = vbRed
    
End Sub

Private Sub mnuSalvar_Click()
    MsgBox ("Nada para salvar")

End Sub
