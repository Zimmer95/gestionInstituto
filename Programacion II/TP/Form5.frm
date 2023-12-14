VERSION 5.00
Begin VB.Form FormProducto 
   Caption         =   "Form5"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12405
   LinkTopic       =   "Form5"
   ScaleHeight     =   5760
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "volver"
      Height          =   495
      Left            =   11040
      TabIndex        =   25
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Borrar campo selec."
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   4680
      Width           =   2055
   End
   Begin VB.ListBox List5 
      Enabled         =   0   'False
      Height          =   2985
      Left            =   10800
      TabIndex        =   18
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Enabled         =   0   'False
      Height          =   2985
      Left            =   9240
      TabIndex        =   17
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      Height          =   2985
      Left            =   7680
      TabIndex        =   16
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   6120
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   4080
      TabIndex        =   14
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2040
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Marca"
      Height          =   375
      Left            =   10800
      TabIndex        =   23
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Stock"
      Height          =   375
      Left            =   9240
      TabIndex        =   22
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Precio c/u"
      Height          =   375
      Left            =   7680
      TabIndex        =   21
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Codigo"
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   600
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   3960
      Y1              =   120
      Y2              =   5880
   End
   Begin VB.Label Label6 
      Caption         =   "Marca"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Stock"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Precio"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre producto"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Nuevo producto"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FormProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub actualizacion()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear

Dim preg As producto
    Open App.Path + "\producto.txt" For Input As 1 Len = Len(preg)
    While Not EOF(1)
        Input #1, preg.producto, preg.codproducto, preg.precio, preg.stock, preg.marca
            List1.AddItem (preg.producto)
            List2.AddItem (preg.codproducto)
            List3.AddItem (preg.precio)
            List4.AddItem (preg.stock)
            List5.AddItem (preg.marca)
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
Text4.Text = ""
    Text4.Enabled = False
Combo1.Enabled = False
    Combo1.Text = ""
    
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Command1.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Ingresar" Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Combo1.Enabled = True
    Text1.SetFocus
    Command1.Caption = "Guardar"
    
Else
    If Command1.Caption = "Guardar" Then
        
            If MsgBox("¿Desea guardar los datos?", vbYesNo + vbCritical + vbDefaultButton2, "Producto") = vbYes Then
                Dim preg As producto
                Open App.Path + "\producto.txt" For Append As 1 Len = Len(preg)
                    Write #1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Combo1.Text
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
    codp = List2.List(r)
    Dim preg As producto
        Open App.Path + "\producto.txt" For Input As 1 Len = Len(preg)
        While Not EOF(1)
            Input #1, preg.producto, preg.codproducto, preg.precio, preg.stock, preg.marca
                If preg.codproducto = codp Then
                    Text1.Text = preg.producto
                        Text1.Enabled = True
                    Text2.Text = preg.codproducto
                        Text2.Enabled = True
                    Text3.Text = preg.precio
                        Text3.Enabled = True
                    Text4.Text = preg.stock
                        Text4.Enabled = True
                    Combo1.Text = preg.marca
                        Combo1.Enabled = True
                End If
        Wend
            Close #1
Else
    Dim dpreg As producto
    Dim dxpreg As producto

    Open App.Path + "\producto.txt" For Input As #1 Len = Len(dpreg)
    Open App.Path + "\paux.txt" For Append As #2 Len = Len(dxpreg)
        While Not EOF(1)
            Input #1, dpreg.producto, dpreg.codproducto, dpreg.precio, dpreg.stock, dpreg.marca
            If dpreg.codproducto = Text2.Text Then
                Write #2, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Combo1.Text
            Else
                Write #2, dpreg.producto, dpreg.codproducto, dpreg.precio, dpreg.stock, dpreg.marca
            End If
        Wend
        Close #1
        Close #2
        
        
    Kill App.Path + "\producto.txt"
    Name App.Path + "\paux.txt" As App.Path + "\producto.txt"
    
    Command2.Caption = "Modificar"
    
    actualizacion
    borrar
End If
End Sub

Private Sub Command3_Click()
borrar
End Sub

Private Sub Command4_Click()
'FormAdmin.Show
'Unload Me


End Sub

Private Sub Command5_Click()
r = List1.ListIndex
codborrar = List2.List(r)
Dim preg As producto
Dim xpreg As producto

    Open App.Path + "\producto.txt" For Input As #1 Len = Len(preg)
    Open App.Path + "\paux.txt" For Append As #2 Len = Len(preg)
        While Not EOF(1)
            Input #1, preg.producto, preg.codproducto, preg.precio, preg.stock, preg.marca
                If codborrar <> preg.codproducto Then
                    Write #2, preg.producto, preg.codproducto, preg.precio, preg.stock, preg.marca
                End If
        Wend
        Close #1
        Close #2
        
    Kill App.Path + "\producto.txt"
    Name App.Path + "\paux.txt" As App.Path + "\producto.txt"



actualizacion

End Sub

Private Sub Command6_Click()
FormAdmin.Show
Unload Me
End Sub

Private Sub Form_Activate()
Command1.SetFocus

Dim creg As categoria
    Open App.Path + "\categoria.txt" For Input As 1 Len = Len(creg)
    While Not EOF(1)
        Input #1, creg.categoria, creg.codcategoria
            Combo1.AddItem (creg.categoria)
    Wend
    Close #1
  
actualizacion

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
xs = 0
Dim preg As producto
    Open App.Path + "\producto.txt" For Input As 1 Len = Len(preg)
    While Not EOF(1)
        Input #1, preg.codproducto
            If preg.codproducto = Text2.Text Then
                xs = 1
            End If
    Wend
    Close #1
    If xs = 0 Then
        Text3.SetFocus
    Else
        r = MsgBox("El codigo del producto ya existe, no se puede repetir", vbOKOnly + 0 + vbDefaultButton1, "Productos")
    End If

End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Combo1.SetFocus
End If
End Sub
