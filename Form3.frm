VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_carreras 
   Caption         =   "Form3"
   ClientHeight    =   8400
   ClientLeft      =   7410
   ClientTop       =   3945
   ClientWidth     =   8700
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   8700
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   7680
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
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
      Connect         =   $"Form3.frx":0000
      OLEDBString     =   $"Form3.frx":00A1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "carrera"
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
      Left            =   360
      TabIndex        =   11
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "Id_carrera"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      DataField       =   "carrera"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "eliminar"
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "guardar"
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "nuevo"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "siguiente"
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ultimo"
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "CARRERA"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   " CARRERAS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      TabIndex        =   8
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frm_carreras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_alumnos_Click()
frm_carreras.Hide
frm_alumnos.Show
End Sub

Private Sub mnu_profesores_Click()
frm_carreras.Hide
frm_profesores.Show
End Sub


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

Adodc1.Recordset.MoveLast
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command5_Click()

Adodc1.Recordset.Delete
End Sub

Private Sub Command6_Click()

Adodc1.Recordset.Update
frm_alumnos.Text1.Refresh
End Sub

Private Sub Command7_Click()
frm_carreras.Hide
frm_menu_principal.Show
End Sub

