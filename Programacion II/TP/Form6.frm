VERSION 5.00
Begin VB.Form FormLocalidad 
   Caption         =   "Form6"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form6"
   ScaleHeight     =   6660
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   3180
      Left            =   7440
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Borrar Campo selecionado"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   5160
      Width           =   3375
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   3180
      Left            =   5760
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "Form6.frx":0000
      Left            =   4080
      List            =   "Form6.frx":0002
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Volver"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Codigo localidad"
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Codigo Postal"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Localidades"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   6600
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo localidad"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Localidad"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo postal"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Agregar localidad"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "FormLocalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub actualizacion()
List1.Clear
List2.Clear
List3.Clear

Dim lreg As localidad
    Open App.Path + "\localidad.txt" For Input As 1 Len = Len(lreg)
    While Not EOF(1)
            Input #1, lreg.localidad, lreg.codpostal, lreg.codlocalidad
                List1.AddItem (lreg.localidad)
                List2.AddItem (lreg.codpostal)
                List3.AddItem (lreg.codlocalidad)
    Wend
    Close #1
End Sub

Private Sub borrar()
Command1.Caption = "Ingresar"
Text1.Text = ""
    Text1.Enabled = False
Text2.Text = ""
    Text2.Enabled = False
Text3.Text = ""
    Text3.Enabled = False
End Sub
Private Sub Command1_Click()
If Command1.Caption = "Ingresar" Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text1.SetFocus
    
Else
    If Command1.Caption = "Guardar" Then
        
            If MsgBox("¿Desea guardar los datos?", vbYesNo + vbCritical + vbDefaultButton2, "Nuevo usuario") = vbYes Then
                Dim reg As localidad
                Open App.Path + "\localidad.txt" For Append As 1 Len = Len(reg)
                    Write #1, Text1.Text, Text2.Text, Text3.Text
                    Close #1
                borrar
                actualizacion
                Command1.Caption = "Ingresar"
            End If
    End If
End If
End Sub

Private Sub Command3_Click()
borrar
Command1.Caption = "Ingresar"
End Sub

Private Sub Command4_Click()
FormAdmin.Show
FormLocalidad.Hide
End Sub

Private Sub Command5_Click()
r = List1.ListIndex
codborrar = List3.List(r)
Dim lreg As localidad
Dim xlreg As localidad

    Open App.Path + "\localidad.txt" For Input As #1 Len = Len(lreg)
    Open App.Path + "\laux.txt" For Append As #2 Len = Len(lreg)
        While Not EOF(1)
            Input #1, lreg.localidad, lreg.codpostal, lreg.codlocalidad
                If codborrar <> lreg.codlocalidad Then
                    Write #2, lreg.localidad, lreg.codpostal, lreg.codlocalidad
                End If
        Wend
        Close #1
        Close #2
        
    Kill App.Path + "\localidad.txt"
    Name App.Path + "\laux.txt" As App.Path + "\localidad.txt"



actualizacion

End Sub

Private Sub Form_Activate()
Command1.SetFocus
actualizacion
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
xs = 0
Dim lreg As localidad
    Open App.Path + "\localidad.txt" For Input As 1 Len = Len(lreg)
    While Not EOF(1)
        Input #1, lreg.localidad
            If lreg.localidad = Text1.Text Then
                xs = 1
            End If
    Wend
    Close #1
    If xs = 0 Then
        Text2.SetFocus
    Else
        r = MsgBox("La Localidad ya existe", vbOKOnly + 0 + vbDefaultButton1, "Nueva localidad")
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
xs = 0
Dim lreg As localidad
    Open App.Path + "\localidad.txt" For Input As 1 Len = Len(lreg)
    While Not EOF(1)
        Input #1, lreg.codlocalidad
            If lreg.codlocalidad = Text3.Text Then
                xs = 1
            End If
    Wend
    Close #1
    If xs = 0 Then
        Command1.SetFocus
        Command1.Caption = "Guardar"
    Else
        r = MsgBox("El codigo de la localidad ya existe, no se puede repetir", vbOKOnly + 0 + vbDefaultButton1, "Nueva localidad")
    End If

End If
End Sub
