VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingreso"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Edad"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese DNI"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Ingreso" Then
    Text1.Enabled = True
    Text1.SetFocus
Else
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Command1.Caption = "Ingreso"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.Enabled = True
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.Enabled = True
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.Enabled = True
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
    Command1.Caption = "Guardar"
End If
End Sub
