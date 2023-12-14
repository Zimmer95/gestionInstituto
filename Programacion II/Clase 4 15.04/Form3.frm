VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4860
   ClientLeft      =   4650
   ClientTop       =   3240
   ClientWidth     =   9015
   LinkTopic       =   "Form3"
   ScaleHeight     =   4860
   ScaleWidth      =   9015
   Begin VB.CommandButton Command1 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese DNI"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Edad"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form3.Hide
End Sub
