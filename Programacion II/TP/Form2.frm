VERSION 5.00
Begin VB.Form FormAdmin 
   Caption         =   "Form2"
   ClientHeight    =   4215
   ClientLeft      =   4080
   ClientTop       =   3510
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   ScaleHeight     =   4215
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar sesion"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Agregar localidad"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar producto"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar categoria"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Admin"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FormAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormMenu.Show
FormAdmin.Hide
End Sub

Private Sub Command2_Click()
FormCategoria.Show
FormAdmin.Hide
End Sub

Private Sub Command3_Click()
FormProducto.Show
FormAdmin.Hide
End Sub

Private Sub Command5_Click()
FormLocalidad.Show
FormAdmin.Hide
End Sub
