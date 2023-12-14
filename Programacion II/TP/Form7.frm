VERSION 5.00
Begin VB.Form FormCompra 
   Caption         =   "Form7"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form7"
   ScaleHeight     =   5475
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Comprar"
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar sesión"
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   5040
      Width           =   1935
   End
   Begin VB.ListBox List3 
      Height          =   2595
      Left            =   7800
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   5160
      TabIndex        =   9
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   3480
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   0
      Y2              =   5520
   End
   Begin VB.Label Label9 
      Caption         =   "Precio"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Nombre de la zapatilla"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Marca"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Nombre de la Zapatilla"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Marca"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Filtros"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Bienvenido"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Tienda de Zapatillas"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "FormCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormMenu.Show
Unload Me
End Sub

Private Sub Form_Activate()
List1.Clear
List2.Clear
List3.Clear

Dim creg As categoria
    Open App.Path + "\categoria.txt" For Input As 1 Len = Len(creg)
    While Not EOF(1)
        Input #1, creg.categoria, creg.codcategoria
            Combo1.AddItem (creg.categoria)
    Wend
    Close #1
  



Dim preg As producto
    Open App.Path + "\producto.txt" For Input As 1 Len = Len(preg)
    While Not EOF(1)
        Input #1, preg.producto, preg.codproducto, preg.precio, preg.stock, preg.marca
            List2.AddItem (preg.producto)
            List3.AddItem (preg.precio)
            List1.AddItem (preg.marca)
            Combo2.AddItem (preg.producto)
    Wend
    Close #1
End Sub
