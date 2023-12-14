VERSION 5.00
Begin VB.Form frm_menu_principal 
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   8100
   ClientTop       =   3780
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   9705
   Begin VB.CommandButton Command4 
      Caption         =   "Profesores"
      Height          =   975
      Left            =   3000
      TabIndex        =   3
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Carreras"
      Height          =   975
      Left            =   3000
      TabIndex        =   2
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Materias"
      Height          =   975
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alumnos"
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
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
