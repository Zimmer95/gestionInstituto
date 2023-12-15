VERSION 5.00
Begin VB.Form frm_login 
   Caption         =   "Login"
   ClientHeight    =   4950
   ClientLeft      =   5025
   ClientTop       =   3120
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   7590
   Begin VB.CommandButton Command1 
      Caption         =   "Inicial sesion"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear usuario"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label3 
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
      Left            =   3120
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "usuario: admin contraseña: admin"
      Height          =   615
      Left            =   6240
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frm_login"
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
        Input #1, ureg.usuario, ureg.contrasenia
            If ureg.usuario = Text1.Text And ureg.contrasenia = Text2.Text Then
                xs = 1
                If ureg.usuario = "admin" And ureg.contrasenia = "admin" Then
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
        MDIForm1.Show
        frm_menu_principal.Show
        frm_menu_principal.SetFocus
        frm_login.Hide
       
End If
    If xs = 2 Then
        Text2.Text = ""
        MDIForm1.Show
        frm_menu_principal.Show
        frm_menu_principal.SetFocus
        frm_login.Hide
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

