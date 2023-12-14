VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "pantalla inicial"
   ClientHeight    =   5415
   ClientLeft      =   4485
   ClientTop       =   2835
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese Nombre:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Form1.Caption = Text1.Text
 Text1.Text = ""
 Text1.Enabled = True
 Text1.SetFocus                 'foco
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'    Text2.Text = KeyAscii
If KeyAscii = 13 Then
    Form1.Caption = Text1.Text
    Text1.Text = ""              'text1 en blanco
    Text1.Enabled = False           'volver false
End If
        'If la pregunta, Then por el si, Else por el no, EndIf fin
End Sub
