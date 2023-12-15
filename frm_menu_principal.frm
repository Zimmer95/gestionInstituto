VERSION 5.00
Begin VB.Form frm_menu_principal 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   4770
   ClientTop       =   1110
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   9705
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Profesores"
      Height          =   975
      Left            =   4800
      TabIndex        =   3
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Carreras"
      Height          =   975
      Left            =   4800
      TabIndex        =   2
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Materias"
      Height          =   975
      Left            =   4800
      TabIndex        =   1
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alumnos"
      Height          =   975
      Left            =   4800
      TabIndex        =   0
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Gestion de instituto"
      Enabled         =   0   'False
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
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000002&
      DrawMode        =   14  'Copy Pen
      Height          =   5415
      Left            =   4560
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000D&
      DrawMode        =   16  'Merge Pen
      Height          =   6495
      Left            =   3960
      Top             =   840
      Width           =   4695
   End
End
Attribute VB_Name = "frm_menu_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frm_menu_principal.Hide
frm_alumnos.Show
End Sub

Private Sub Command2_Click()
frm_menu_principal.Hide
frm_materias.Show
End Sub

Private Sub Command3_Click()
frm_menu_principal.Hide
frm_carreras.Show
End Sub

Private Sub Command4_Click()
frm_menu_principal.Hide
frm_profesores.Show
End Sub

