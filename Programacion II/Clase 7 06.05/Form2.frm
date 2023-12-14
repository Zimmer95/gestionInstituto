VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5985
   ClientLeft      =   4485
   ClientTop       =   3255
   ClientWidth     =   9225
   LinkTopic       =   "Form2"
   ScaleHeight     =   5985
   ScaleWidth      =   9225
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3720
      TabIndex        =   17
      Top             =   2640
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trabaja"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   840
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Si"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   2040
      List            =   "Form2.frx":000D
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingreso"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4560
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
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Ingrese nueva localidad"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Localidad"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5520
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
Combo1.Text = ""
    Combo1.Enabled = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Combo1.Text <> Empty Then
    Frame1.Enabled = True
    Option1.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Ingreso" Then
    Text1.Enabled = True
    Text1.SetFocus
Else
    If MsgBox("¿Desea guardar los datos?", vbYesNo + vbCritical + vbDefaultButton2, "Alta de cliente") = vbYes Then
        Form3.Label5.Caption = Text1.Text
        Form3.Label6.Caption = Text2.Text
        Form3.Label7.Caption = Text3.Text
        Form3.Label8.Caption = Text4.Text
        'Form3.List1.AddItem (Text1.Text + "-" + Text2.Text + "-" + Text3.Text + "-" + Text4.Text + "-" + Combo1.Text)
        Form3.List1.AddItem ("Datos del Cliente")
        Form3.List1.AddItem ("Nombre:" + Text3.Text + "__" + "Apellido:" + Text2.Text)
        Form3.List1.AddItem ("")
        Form3.List1.AddItem ("Orundo de la localidad de:" + Combo1.Text)
            If Option1.Value = True Then
                Form3.List1.AddItem ("Usuario Trabaja")
                Form3.Option1 = True
            Else
                Form3.List1.AddItem ("Usuario No trabaja")
                Form3.Option2 = True
            End If
        Form3.List1.AddItem ("_____________________________")
        Form3.Show
        Form2.Hide
        limpia
    End If

End If
End Sub

Private Sub Command2_Click()
limpia
End Sub

Private Sub Command3_Click()
Form2.Hide
End Sub

Private Sub Command4_Click()
    If MsgBox("¿Desea agregar nueva localidad?", vbYesNo + vbCritical + vbDefaultButton2, "Alta de cliente") = vbYes Then
        Text5.Visible = True
        Label7.Visible = True
        Command5.Visible = True
   End If
End Sub

Private Sub Command5_Click()
    If MsgBox("¿Desea guardar nueva localidad?", vbYesNo + vbCritical + vbDefaultButton2, "Alta de cliente") = vbYes Then
        Combo1.AddItem (Text5.Text)
        Text5.Visible = False
        Label7.Visible = False
        Command5.Visible = False
    Else
        Text5.Visible = False
        Label7.Visible = False
        Command5.Visible = False
        End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = ""
End Sub

Private Sub Option1_Click()
Command1.Caption = "Guardar"
    Command1.SetFocus
End Sub

Private Sub Option2_Click()
Command1.Caption = "Guardar"
    Command1.SetFocus
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
    Combo1.Enabled = True
    Combo1.SetFocus
End If
End Sub

Private Sub Text4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Ingrese su Edad"
End Sub

