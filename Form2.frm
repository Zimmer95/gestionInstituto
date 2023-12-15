VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_profesores 
   Caption         =   "PROFESORES"
   ClientHeight    =   8565
   ClientLeft      =   5400
   ClientTop       =   1965
   ClientWidth     =   8925
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   8925
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3600
      Top             =   8280
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":00A1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "docente"
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
      Left            =   240
      TabIndex        =   17
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text5 
      DataField       =   "Id_docente"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   1320
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   615
      Left            =   2520
      TabIndex        =   13
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ultimo"
      Height          =   615
      Left            =   6120
      TabIndex        =   12
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "siguiente"
      Height          =   615
      Left            =   4320
      TabIndex        =   11
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "nuevo"
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "guardar"
      Height          =   615
      Left            =   6120
      TabIndex        =   9
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "eliminar"
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      DataField       =   "materia"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   4440
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      DataField       =   "matricula"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3720
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      DataField       =   "apellido"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   "ID"
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   " PROFESORES"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "MATERIA"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "MATRICULA"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "APELLIDO"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "NOMBRE"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "frm_profesores"
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
Adodc1.Recordset.MoveNext

If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext

If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
Form1.Text1.Refresh
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub mnu_alumnos_Click()
frm_profesores.Hide
frm_alumnos.Show
End Sub

Private Sub mnu_carreras_Click()
frm_profesores.Hide
frm_carreras.Show
End Sub

Private Sub Command7_Click()
frm_profesores.Hide
frm_menu_principal.Show
End Sub

