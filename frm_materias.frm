VERSION 5.00
Begin VB.Form frm_materias 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      DataField       =   "apellido"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      DataField       =   "dni"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox Text4 
      DataField       =   "direccion"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ultimo"
      Height          =   615
      Left            =   5880
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "siguiente"
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "nuevo"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "guardar"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      DataField       =   "Id_alumno"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<- Atras"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "NOMBRE"
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "APELLIDO"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "MATERIAS"
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "DNI"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "DIRECCION"
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "ID"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frm_materias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
