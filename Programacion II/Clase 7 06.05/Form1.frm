VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "Pantalla incial"
   ClientHeight    =   4560
   ClientLeft      =   555
   ClientTop       =   840
   ClientWidth     =   6315
   FillColor       =   &H80000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese nombre"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1695
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
 Text1.SetFocus
 End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Text2.Text = KeyAscii
 If KeyAscii = 13 Then
    Form1.Caption = Text1.Text
    Text1.Text = ""
    Text1.Enabled = False
    
End If
    Text2.Text = KeyAscii
    
 
End Sub

