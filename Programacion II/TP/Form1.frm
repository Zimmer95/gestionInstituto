VERSION 5.00
Begin VB.Form FormMenu 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear usuario"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inicial sesion"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Nota: Cuenta administrativa es cuenta: admin pass: admin"
      Height          =   975
      Left            =   4080
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Principal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
xs = 0
'adm = 0
Dim ureg As usuario
    Open App.Path + "\usuario.txt" For Input As 1 Len = Len(ureg)
    While Not EOF(1)
        Input #1, ureg.usuario, ureg.apellido, ureg.pass, ureg.localidad
            If ureg.usuario = Text1.Text And ureg.pass = Text2.Text Then
                xs = 1
                If ureg.usuario = "admin" And ureg.pass = "admin" Then
                    xs = 2
                    
                   'adm = 1
                End If
            End If
    Wend
    Close #1
    If xs = 0 Then
         r = MsgBox("El usuario y contraseña no son correctos", vbOKOnly + 0 + vbDefaultButton1, "Inicir sesion")
        Text2.Text = ""
End If
    If xs = 1 Then
        Text2.Text = ""
        FormCompra.Show
        FormCompra.Label3.Caption = Text1.Text
        FormMenu.Hide
       
End If
    If xs = 2 Then
        Text2.Text = ""
        FormAdmin.Show
        FormMenu.Hide
End If
    
End Sub

Private Sub Command2_Click()
FormNuevoUsua.Show
FormMenu.Hide
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Command1.SetFocus
End If
End Sub
