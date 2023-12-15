VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_materias 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   5610
   ClientTop       =   2235
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   9075
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
      Top             =   7800
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
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
      Connect         =   $"frm_materias.frx":0000
      OLEDBString     =   $"frm_materias.frx":00A1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "materia"
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
   Begin VB.TextBox Text1 
      DataField       =   "materia"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      DataField       =   "cargaHoraria"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      DataField       =   "id_carrera"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ultimo"
      Height          =   615
      Left            =   6480
      TabIndex        =   6
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "siguiente"
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "nuevo"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "guardar"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      DataField       =   "Id_materia"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<- Atras"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "MATERIA"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "CARGA HORARIA"
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "MATERIAS"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "ID DE CARRERA"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "ID"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frm_materias"
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
End Sub

Private Sub Command5_Click()

Adodc1.Recordset.Update
frm_alumnos.Text1.Refresh
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command7_Click()
frm_materias.Hide
frm_menu_principal.Show
End Sub

