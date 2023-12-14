VERSION 5.00
Begin VB.Form FormNuevoUsua 
   Caption         =   "Form3"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form3"
   ScaleHeight     =   4680
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form3.frx":0000
      Left            =   1800
      List            =   "Form3.frx":0002
      TabIndex        =   9
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Volver"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar datos"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Contraseña"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nuevo usuario"
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
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Localidad"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre y Apellido"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "FormNuevoUsua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub borrar()
Command1.Caption = "Ingresar"
Text1.Text = ""
    Text1.Enabled = False
Text2.Text = ""
    Text2.Enabled = False
Text3.Text = ""
    Text2.Enabled = False
Combo1.Text = ""
    Combo1.Enabled = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Command1.Caption = "Guardar"
    Command1.SetFocus
    
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Ingresar" Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Combo1.Enabled = True
    Text1.SetFocus
Else
    If Command1.Caption = "Guardar" Then
        If MsgBox("¿Desea registrarte?", vbYesNo + vbCritical + vbDefaultButton2, "Nuevo usuario") = vbYes Then
            Dim reg As usuario
            Open App.Path + "\usuario.txt" For Append As 1 Len = Len(reg)
                Write #1, Text1.Text, Text2.Text, Text3.Text, Combo1.Text
                Close #1
            borrar
            FormMenu.Show
            FormNuevoUsua.Hide
        End If
    
    End If
End If

End Sub

Private Sub Command3_Click()
borrar
End Sub

Private Sub Command4_Click()
FormMenu.Show
FormNuevoUsua.Hide
End Sub

Private Sub Form_Activate()
Command1.SetFocus

Dim lreg As localidad
    Open App.Path + "\localidad.txt" For Input As 1 Len = Len(lreg)
    While Not EOF(1)
        Input #1, lreg.localidad, lreg.codpostal, lreg.codlocalidad
            Combo1.AddItem (lreg.localidad)
    Wend
    Close #1
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
xs = 0
Dim ureg As usuario
    Open App.Path + "\usuario.txt" For Input As 1 Len = Len(ureg)
    While Not EOF(1)
        Input #1, ureg.usuario
            If ureg.usuario = Text1.Text Then
                xs = 1
            End If
    Wend
    Close #1
    If xs = 0 Then
        Text2.SetFocus
    Else
        r = MsgBox("El usuario ya existe, prueba con otro nombre.", vbOKOnly + 0 + vbDefaultButton1, "Nueva Usuario")
    End If

End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_Change()
If KeyAscii = 13 And Text1.Text <> Empty Then
    Combo1.SetFocus
End If
End Sub
