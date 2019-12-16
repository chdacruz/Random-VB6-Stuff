VERSION 5.00
Begin VB.Form frmLista 
   Caption         =   "Lista"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox listNomes 
      Height          =   1425
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Limpar Tudo"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remover"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Adicionar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    listNomes.AddItem (txtNome.Text)
    txtNome = ""
    
End Sub

Private Sub cmdClear_Click()
    listNomes.Clear
    
End Sub

Private Sub cmdRemove_Click()
    If listNomes.ListIndex < 0 Then
        MsgBox ("Nenhum item selecionado")
    Else
        listNomes.RemoveItem (listNomes.ListIndex)
    End If
    
End Sub

Private Sub txtNome_Change()
    If txtNome.Text = "" Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
    
End Sub
