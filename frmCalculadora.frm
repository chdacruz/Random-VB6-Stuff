VERSION 5.00
Begin VB.Form frmCalculadora 
   Caption         =   "Calculadora"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame1 
      Caption         =   "Painel"
      Height          =   3975
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   5655
      Begin VB.CommandButton cmdIgual 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4560
         TabIndex        =   18
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4560
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdMais 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   16
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdMenos 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   15
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdMult 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdDiv 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdVirgula 
         Caption         =   ","
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   12
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmd0 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton cmd3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   10
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   9
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmd6 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   7
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmd5 
         Caption         =   "5"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmd4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmd9 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmd8 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmd7 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtCalc 
      Height          =   735
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variavél usada para verificar o começo de uma nova operação de modo a limpar o TextField
Dim isDone As Boolean


Function Sum(n1 As Double, n2 As Double) As Double
    Sum = n1 + n2
End Function

Function Subt(n1 As Double, n2 As Double) As Double
    Subt = n1 - n2
End Function

Function Times(n1 As Double, n2 As Double) As Double
    Times = n1 * n2
End Function

Function Div(n1 As Double, n2 As Double) As Double
    Div = n1 / n2
End Function

Function OpsController(formula() As String)
    Dim numbers3(100) As Double
    Dim i As Integer
    
    Dim total As Double
    total = 0
    
    'Convertendo para Double e armazenando no rolê
    For i = 0 To UBound(formula())
        'Se for número
        If (IsNumeric(formula(i))) Then
            numbers3(i) = CDbl(formula(i))
        Else
            ''Desta forma, ambas as arrays formula() quanto a numbers3() terão o mesmo tamanho. Onde houver símbolos em formula,
            'haverá o número 0 em numbers3()
            numbers3(i) = 0
        End If
    Next
    
    'Agora comparar numbers() com formula() para fazer as operações na ordem correta
    For i = 0 To UBound(formula())
        'Se for um símbolo, realizar operação
        If (IsNumeric(formula(i)) = False) Then
            If (formula(i) = "+") Then
                total = Sum(numbers3(i - 1), numbers3(i + 1))
                ''Zerar n1 após a operação (Pop), colocar o total da operação em n2 (Push)
                numbers3(i - 1) = 0
                numbers3(i + 1) = total
            ElseIf (formula(i) = "-") Then
                total = Subt(numbers3(i - 1), numbers3(i + 1))
                ''Zerar n1 após a operação (Pop), colocar o total da operação em n2 (Push)
                numbers3(i - 1) = 0
                numbers3(i + 1) = total
            ElseIf (formula(i) = "*") Then
                total = Times(numbers3(i - 1), numbers3(i + 1))
                ''Zerar n1 após a operação (Pop), colocar o total da operação em n2 (Push)
                numbers3(i - 1) = 0
                numbers3(i + 1) = total
            ElseIf (formula(i) = "/") Then
                total = Div(numbers3(i - 1), numbers3(i + 1))
                ''Zerar n1 após a operação (Pop), colocar o total da operação em n2 (Push)
                numbers3(i - 1) = 0
                numbers3(i + 1) = total
            End If
        End If
    Next
    
    Me.txtCalc = CStr(total)
    'Sinaliza o final da operação
    isDone = True
    
End Function

Function StrToArray(formula As String) As String()
    Dim numbers4(100) As String
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To Len(formula)
        'Para fazer com que números quebrados possam ser processados como double
        j = i + 1
        If (Mid(formula, i + 1, 1) = ",") Then
            Do While IsNumeric(Mid(formula, i + 1, 1))
                j = j + 1
                Mid(formula, i + 1, 1) = Mid(formula, i + 1, 1) & Mid(formula, j, 1)
                'formula(i) = formula(i) & formula(j)
            Loop
        End If
        numbers4(i) = Mid(formula, i + 1, 1)
    Next
    
    StrToArray = numbers4()

End Function

Function isCalculationDone()
    If (isDone = True) Then
        Me.txtCalc = ""
        isDone = False
    End If
End Function



Private Sub cmd0_Click()
    Me.txtCalc = Me.txtCalc & "0"
End Sub

Private Sub cmd1_Click()
    Me.txtCalc = Me.txtCalc & "1"
End Sub

Private Sub cmd2_Click()
    Me.txtCalc = Me.txtCalc & "2"
End Sub

Private Sub cmd3_Click()
    Me.txtCalc = Me.txtCalc & "3"
End Sub

Private Sub cmd4_Click()
    Me.txtCalc = Me.txtCalc & "4"
End Sub

Private Sub cmd5_Click()
    Me.txtCalc = Me.txtCalc & "5"
End Sub

Private Sub cmd6_Click()
    Me.txtCalc = Me.txtCalc & "6"
End Sub

Private Sub cmd7_Click()
    Me.txtCalc = Me.txtCalc & "7"
End Sub

Private Sub cmd8_Click()
    Me.txtCalc = Me.txtCalc & "8"
End Sub

Private Sub cmd9_Click()
    Me.txtCalc = Me.txtCalc & "9"
End Sub

Private Sub cmdClear_Click()
    Me.txtCalc = ""
End Sub

Private Sub cmdDiv_Click()
    Me.txtCalc = Me.txtCalc & "/"
End Sub

Private Sub cmdIgual_Click()
    Call OpsController(StrToArray(Me.txtCalc))
End Sub

Private Sub cmdMais_Click()
    Me.txtCalc = Me.txtCalc & "+"
End Sub

Private Sub cmdMenos_Click()
    Me.txtCalc = Me.txtCalc & "-"
End Sub

Private Sub cmdMult_Click()
    Me.txtCalc = Me.txtCalc & "*"
End Sub

Private Sub cmdVirgula_Click()
    Me.txtCalc = Me.txtCalc & ","
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call isCalculationDone

    Select Case Chr(KeyAscii)
        Case "0"
            Me.txtCalc = Me.txtCalc & "0"
        Case "1"
            Me.txtCalc = Me.txtCalc & "1"
        Case "2"
            Me.txtCalc = Me.txtCalc & "2"
        Case "3"
            Me.txtCalc = Me.txtCalc & "3"
        Case "4"
            Me.txtCalc = Me.txtCalc & "4"
        Case "5"
            Me.txtCalc = Me.txtCalc & "5"
        Case "6"
            Me.txtCalc = Me.txtCalc & "6"
        Case "7"
            Me.txtCalc = Me.txtCalc & "7"
        Case "8"
            Me.txtCalc = Me.txtCalc & "8"
        Case "9"
            Me.txtCalc = Me.txtCalc & "9"
        Case ","
            Me.txtCalc = Me.txtCalc & ","
        Case "/"
            Me.txtCalc = Me.txtCalc & "/"
        Case "*"
            Me.txtCalc = Me.txtCalc & "*"
        Case "+"
            Me.txtCalc = Me.txtCalc & "+"
        Case "-"
            Me.txtCalc = Me.txtCalc & "-"
        Case ","
            Me.txtCalc = Me.txtCalc & ","
        Case "="
            Call OpsController(StrToArray(Me.txtCalc))
        'Para a tecla enter
        Case Chr(13)
           Call OpsController(StrToArray(Me.txtCalc))
        Case "c"
            Me.txtCalc = ""
            
    End Select
End Sub

