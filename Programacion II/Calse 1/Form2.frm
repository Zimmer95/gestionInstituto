VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4575
   ClientLeft      =   4065
   ClientTop       =   3045
   ClientWidth     =   10800
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   10800
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingreso"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Edad"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese DNI"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Ingreso" Then
 Text1.Enabled = True
    Text1.SetFocus
 Command1.Caption = "Guardar"
Else
         Command1.Caption = "Guardar"
         Text1.Text = ""
           Text1.Enabled = False
         Text2.Text = ""
           Text2.Enabled = False
         Text3.Text = ""
          Text3.Enabled = False
         Text4.Text = ""
          Text4.Enabled = False
         Command1.Caption = "Ingreso"
    End If
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
End If
End Sub
