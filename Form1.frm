VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_alumnos 
   Caption         =   "ALUMNOS"
   ClientHeight    =   8430
   ClientLeft      =   7920
   ClientTop       =   2550
   ClientWidth     =   9015
   LinkTopic       =   "frm_alumnos"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   9015
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
      Top             =   8160
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":00A1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "alumno"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<- Atras"
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Text5 
      DataField       =   "Id_alumno"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   15
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   615
      Left            =   5880
      TabIndex        =   14
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "guardar"
      Height          =   615
      Left            =   7680
      TabIndex        =   13
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "nuevo"
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "siguiente"
      Height          =   615
      Left            =   5880
      TabIndex        =   11
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ultimo"
      Height          =   615
      Left            =   7680
      TabIndex        =   10
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   615
      Left            =   4080
      TabIndex        =   9
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      DataField       =   "direccion"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   3960
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      DataField       =   "dni"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   3240
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      DataField       =   "apellido"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label6 
      Caption         =   "ID"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "DIRECCION"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "DNI"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "ALUMNOS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "APELLIDO"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "NOMBRE"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "frm_alumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.MovePrevious

If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveNext
End If

End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext

If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.AddNew
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
frm_alumnos.Text1.Refresh
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Delete
frm_alumnos.Text1.Refresh
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Adodc1.Recordset.MoveLast
End Sub


Private Sub mnu_carreras_Click()
frm_alumnos.Hide
frm_carreras.Show
End Sub

Private Sub mnu_profesores_Click()
frm_alumnos.Hide
frm_profesores.Show
End Sub

Private Sub Command7_Click()
frm_alumnos.Hide
frm_menu_principal.Show
End Sub

Private Sub Form_Load()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Text2.Enabled = True
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> Empty Then
    Text3.Enabled = True
    Text3.SetFocus
End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text3.Text <> Empty Then
    Text4.Enabled = True
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text4.Text <> Empty Then
    Command5.SetFocus
End If
End Sub
