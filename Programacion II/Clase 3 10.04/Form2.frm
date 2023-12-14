VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4980
   ClientLeft      =   4485
   ClientTop       =   3255
   ClientWidth     =   9225
   LinkTopic       =   "Form2"
   ScaleHeight     =   4980
   ScaleWidth      =   9225
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingreso"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4560
      Width           =   9135
   End
   Begin VB.Label Label4 
      Caption         =   "Edad"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese DNI"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub limpia()
Command1.Caption = "Ingreso"
Text1.Text = ""
Text1.Enabled = False
Text2.Text = ""
Text2.Enabled = False
Text3.Text = ""
Text3.Enabled = False
Text4.Text = ""
Text4.Enabled = False
End Sub
Private Sub Command1_Click()
If Command1.Caption = "Ingreso" Then
    Text1.Enabled = True
    Text1.SetFocus
Else
    limpia
    MsgBox ("Los datos se grabaron correctamente")
End If
End Sub

Private Sub Command2_Click()
limpia
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And Text1.Text <> Empty Then   'emty=no tenga nada
    Text2.Enabled = True
    Text2.SetFocus
End If
    
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Ingrese Nro. de DNI"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> Empty Then
    Text3.Enabled = True
    Text3.SetFocus
End If
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Ingrese su Apellido"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text3.Text <> Empty Then
    Text4.Enabled = True
    Text4.SetFocus
End If
End Sub


Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Ingrese su Nombre"
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text4.Text <> Empty Then
    Command1.Caption = "Guardar"
    Command1.SetFocus
End If
End Sub

Private Sub Text4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Ingrese su Edad"
End Sub
