VERSION 5.00
Begin VB.Form FormCategoria 
   Caption         =   "Form4"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form4"
   ScaleHeight     =   5130
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Borrar categoria selecionada"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   3720
      Width           =   3255
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   7080
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   4680
      TabIndex        =   9
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Volver"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Marcas registradas"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   4560
      Y1              =   0
      Y2              =   5040
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Marca"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Nueva Marca de Zapatilla"
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
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "FormCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub actualizacion()
List1.Clear
List2.Clear
Dim creg As categoria
    Open App.Path + "\categoria.txt" For Input As 1 Len = Len(creg)
    While Not EOF(1)
            Input #1, creg.categoria, creg.codcategoria
                List1.AddItem (creg.categoria)
                List2.AddItem (creg.codcategoria)
    Wend
    Close #1
End Sub

Private Sub borrar()
Command1.Caption = "Ingresar"
Text1.Text = ""
    Text1.Enabled = False
Text2.Text = ""
    Text2.Enabled = False
End Sub
Private Sub Command1_Click()
If Command1.Caption = "Ingresar" Then
    Text1.Enabled = True
    Text1.SetFocus
Else
    If Command1.Caption = "Guardar" Then
        
            If MsgBox("¿Desea guardar los datos?", vbYesNo + vbCritical + vbDefaultButton2, "Nuevo usuario") = vbYes Then
                Dim creg As categoria
                Open App.Path + "\categoria.txt" For Append As 1 Len = Len(creg)
                    Write #1, Text1.Text, Text2.Text
                    Close #1
                borrar
                actualizacion
                Command1.Caption = "Ingresar"
            End If
    End If
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Modificar" Then
    Command2.Caption = "Guardar Modificaciones"
    r = List1.ListIndex
    codc = List2.List(r)
    Dim creg As categoria
        Open App.Path + "\categoria.txt" For Input As 1 Len = Len(creg)
        While Not EOF(1)
            Input #1, creg.categoria, creg.codcategoria
                If creg.codcategoria = codc Then
                    Text1.Text = creg.categoria
                        Text1.Enabled = True
                    Text2.Text = creg.codcategoria
                        Text2.Enabled = True
                End If
        Wend
            Close #1
Else
    Dim cgreg As categoria
    Dim xcgreg As categoria

    Open App.Path + "\categoria.txt" For Input As #1 Len = Len(cgreg)
    Open App.Path + "\caux.txt" For Append As #2 Len = Len(xcgreg)
        While Not EOF(1)
            Input #1, creg.categoria, creg.codcategoria
            If creg.codcategoria = Text2.Text Then
                Write #2, Text1.Text, Text2.Text
            Else
                Write #2, creg.categoria, creg.codcategoria
            End If
        Wend
        Close #1
        Close #2
        
        
    Kill App.Path + "\categoria.txt"
    Name App.Path + "\caux.txt" As App.Path + "\categoria.txt"
    
    Command2.Caption = "Modificar"
    
    actualizacion
    borrar
End If
End Sub

Private Sub Command4_Click()
FormAdmin.Show
FormCategoria.Hide
End Sub

Private Sub Command5_Click()
r = List1.ListIndex
codborrar = List2.List(r)
Dim creg As categoria
Dim xcreg As categoria

    Open App.Path + "\categoria.txt" For Input As #1 Len = Len(creg)
    Open App.Path + "\caux.txt" For Append As #2 Len = Len(creg)
        While Not EOF(1)
            Input #1, creg.categoria, creg.codcategoria
                If codborrar <> creg.codcategoria Then
                    Write #2, creg.categoria, creg.codcategoria
                End If
        Wend
        Close #1
        Close #2
        
    Kill App.Path + "\categoria.txt"
    Name App.Path + "\caux.txt" As App.Path + "\categoria.txt"
    
actualizacion

End Sub

Private Sub Form_Activate()
Command1.SetFocus
actualizacion
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
xs = 0
Dim creg As categoria
    Open App.Path + "\categoria.txt" For Input As 1 Len = Len(creg)
    While Not EOF(1)
        Input #1, creg.categoria
            If creg.categoria = Text1.Text Then
                xs = 1
            End If
    Wend
    Close #1
    If xs = 0 Then
        Text2.Enabled = True
        Text2.SetFocus
    Else
        r = MsgBox("La Marca ya existe", vbOKOnly + 0 + vbDefaultButton1, "Categorias")
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
xs = 0
Dim creg As categoria
    Open App.Path + "\categoria.txt" For Input As 1 Len = Len(creg)
    While Not EOF(1)
        Input #1, creg.codcategoria
            If creg.codcategoria = Text2.Text Then
                xs = 1
            End If
    Wend
    Close #1
    If xs = 0 Then
        Command1.Caption = "Guardar"
        Command1.SetFocus
    Else
        r = MsgBox("El codigo ya existe", vbOKOnly + 0 + vbDefaultButton1, "Categorias")
    End If
End If
End Sub
