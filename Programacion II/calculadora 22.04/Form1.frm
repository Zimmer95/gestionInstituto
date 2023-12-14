VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculadora"
   ClientHeight    =   5175
   ClientLeft      =   2250
   ClientTop       =   2685
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   4215
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command18 
      Caption         =   "<--"
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "C"
      Height          =   495
      Left            =   1800
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton CommandIgual 
      Caption         =   "="
      Height          =   1095
      Left            =   3240
      TabIndex        =   16
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton CommandDivi 
      Caption         =   "/"
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton CommandMulti 
      Caption         =   "x"
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton CommandMenos 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton CommandMas 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton CommandPunto 
      Caption         =   "."
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   3960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   4080
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   240
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   4080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numero1, numero2 As Double
Dim general As String

Private Sub Command1_Click()
Text1.Text = Text1.Text & "1"
End Sub

Private Sub Command10_Click()
Text1.Text = Text1.Text & "0"
End Sub

Private Sub Command17_Click()
Text1.Text = ""
numero1 = 0
numero2 = 0
End Sub

Private Sub Command18_Click()
Text1.Text = StrReverse(Mid(StrReverse(Text1.Text), 2))
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text & "2"
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text & "3"
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text & "4"
End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text & "5"
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text & "6"
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text & "7"
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text & "8"
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text & "9"
End Sub

Private Sub CommandDivi_Click()
numero1 = Text1.Text
Text1.Text = ""
general = "/"
End Sub

Private Sub CommandIgual_Click()
numero2 = Text1.Text
If general = "+" Then
    Text1.Text = numero1 + numero2
    numero1 = Text1.Text
   
Else
If general = "-" Then
    Text1.Text = numero1 - numero2
    numero1 = Text1.Text
   
Else
If general = "X" Then
    Text1.Text = numero1 * numero2
    numero1 = Text1.Text
    
Else
If general = "/" Then
    Text1.Text = numero1 / numero2
    numero1 = Text1.Text
    
End If
End If
End If
End If
End Sub

Private Sub CommandMas_Click()
numero1 = Text1.Text + numero1
Text1.Text = ""
general = "+"
End Sub

Private Sub CommandMas_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
numero1 = Text1.Text + numero1
Text1.Text = ""
general = "+"
End If
End Sub

Private Sub CommandMenos_Click()
numero1 = Text1.Text - numero1
Text1.Text = ""
general = "-"
End Sub

Private Sub CommandMulti_Click()
numero1 = Text1.Text * numero1
Text1.Text = ""
general = "X"
End Sub

Private Sub CommandPunto_Click()
Text1.Text = Text1.Text & "."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
    numero1 = Text1.Text + numero1
    general = "+"
    Text1.Text = ""
    Text2.SetFocus
End If
If KeyAscii = 45 Then
    numero1 = Text1.Text - numero1
    Text1.Text = ""
    general = "-"
    Text2.SetFocus
End If
If KeyAscii = 42 Then
    numero1 = Text1.Text * numero1
    Text1.Text = ""
    general = "X"
    Text2.SetFocus
End If
If KeyAscii = 47 Then
    numero1 = Text1.Text
    Text1.Text = ""
    general = "/"
    Text2.SetFocus
End If
If KeyAscii = 13 Then
    numero2 = Text1.Text

If general = "+" Then
    Text1.Text = numero1 + numero2
    numero1 = Text1.Text
   
Else
If general = "-" Then
    Text1.Text = numero1 - numero2
    numero1 = Text1.Text
   
Else
If general = "X" Then
    Text1.Text = numero1 * numero2
    numero1 = Text1.Text
    
Else
If general = "/" Then
    Text1.Text = numero1 / numero2
    numero1 = Text1.Text
End If
End If
End If
End If
End If
If KeyAscii = 8 Then
    Text1.Text = StrReverse(Mid(StrReverse(Text1.Text), 2))
End If
If KeyAscii = 99 Then
    Text1.Text = ""
    numero1 = 0
    numero2 = 0
     Text2.SetFocus
End If
End Sub

Private Sub Text2_GotFocus()
Text1.Text = Text2.Text
Text1.SetFocus
End Sub
